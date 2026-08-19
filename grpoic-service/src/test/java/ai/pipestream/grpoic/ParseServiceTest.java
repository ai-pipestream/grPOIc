package ai.pipestream.grpoic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ai.pipestream.grpoic.server.PoiParseServiceImpl;
import ai.pipestream.poi.v1.DocumentFormat;
import ai.pipestream.poi.v1.GetServiceInfoRequest;
import ai.pipestream.poi.v1.GetServiceInfoResponse;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseRequestChunk;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.PoiParseServiceGrpc;
import ai.pipestream.poi.v1.SheetCell;
import com.google.protobuf.ByteString;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.StreamObserver;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * Round trips through the real gRPC service (in-process transport) using
 * fixtures POI itself authors in memory, so no binary files are committed and
 * every assertion is against content this test placed.
 */
class ParseServiceTest {

  private static final long CAP_BYTES = 4L * 1024 * 1024;
  private static Server server;
  private static ManagedChannel channel;
  private static ExecutorService executor;

  @BeforeAll
  static void startServer() throws Exception {
    executor = Executors.newVirtualThreadPerTaskExecutor();
    String name = InProcessServerBuilder.generateName();
    server = InProcessServerBuilder.forName(name).directExecutor()
        .addService(new PoiParseServiceImpl(CAP_BYTES, 4, executor)).build().start();
    channel = InProcessChannelBuilder.forName(name).directExecutor().build();
  }

  @AfterAll
  static void stopServer() throws Exception {
    channel.shutdownNow();
    server.shutdownNow();
    executor.shutdown();
    assertTrue(channel.awaitTermination(5, TimeUnit.SECONDS), "channel drain");
  }

  // --- fixtures -----------------------------------------------------------

  private static byte[] docxFixture() throws Exception {
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.getProperties().getCoreProperties().setTitle("Quarterly Narrative");
      document.getProperties().getCoreProperties().setCreator("Archivist");
      document.getProperties().getCustomProperties().addProperty("approved", true);
      var heading = document.createParagraph();
      heading.setStyle("Heading1");
      heading.createRun().setText("Introduction");
      document.createParagraph().createRun().setText("Numbers improved across the board.");
      XWPFTable table = document.createTable(2, 2);
      table.getRow(0).getCell(0).setText("Region");
      table.getRow(0).getCell(1).setText("Revenue");
      table.getRow(1).getCell(0).setText("East");
      table.getRow(1).getCell(1).setText("1200");
      document.write(out);
      return out.toByteArray();
    }
  }

  private static byte[] xlsxFixture() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet("Data");
      var header = sheet.createRow(0);
      header.createCell(0).setCellValue("amount");
      header.createCell(1).setCellValue("active");
      var row = sheet.createRow(1);
      row.createCell(0).setCellValue(42.5);
      row.createCell(1).setCellValue(true);
      var formulaCell = row.createCell(2);
      formulaCell.setCellFormula("A2*2");
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      workbook.createSheet("Empty");
      workbook.write(out);
      return out.toByteArray();
    }
  }

  private static byte[] pptxFixture() throws Exception {
    try (XMLSlideShow show = new XMLSlideShow();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var slide = show.createSlide();
      XSLFTextBox box = slide.createTextBox();
      box.setAnchor(new java.awt.geom.Rectangle2D.Double(50, 50, 400, 100));
      box.setText("The archivist's dream machine");
      show.write(out);
      return out.toByteArray();
    }
  }

  private static byte[] xlsFixture() throws Exception {
    try (HSSFWorkbook workbook = new HSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet("Legacy");
      sheet.createRow(0).createCell(0).setCellValue("still works");
      workbook.write(out);
      return out.toByteArray();
    }
  }

  private static byte[] pptFixture() throws Exception {
    try (HSLFSlideShow show = new HSLFSlideShow();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var slide = show.createSlide();
      HSLFTextBox box = slide.createTextBox();
      box.setText("legacy slide text");
      show.write(out);
      return out.toByteArray();
    }
  }

  // --- round-trip helper --------------------------------------------------

  private record ParseResult(List<ParseEvent> events, Throwable error) {
    ParseStatus status() {
      ParseEvent last = events.get(events.size() - 1);
      assertTrue(last.hasStatus(), "last event must be the status");
      return last.getStatus();
    }
  }

  private static ParseResult parse(byte[] bytes, String documentId, int chunkSize)
      throws InterruptedException {
    List<ParseEvent> events = new ArrayList<>();
    AtomicReference<Throwable> failure = new AtomicReference<>();
    CountDownLatch done = new CountDownLatch(1);
    StreamObserver<ParseRequestChunk> requests =
        PoiParseServiceGrpc.newStub(channel).parseDocument(new StreamObserver<>() {
          @Override
          public void onNext(ParseEvent event) {
            synchronized (events) {
              events.add(event);
            }
          }

          @Override
          public void onError(Throwable error) {
            failure.set(error);
            done.countDown();
          }

          @Override
          public void onCompleted() {
            done.countDown();
          }
        });
    for (int offset = 0; offset < bytes.length; offset += chunkSize) {
      int end = Math.min(bytes.length, offset + chunkSize);
      ParseRequestChunk.Builder chunk = ParseRequestChunk.newBuilder()
          .setData(ByteString.copyFrom(bytes, offset, end - offset))
          .setComplete(end == bytes.length);
      if (offset == 0) chunk.setDocumentId(documentId);
      requests.onNext(chunk.build());
    }
    requests.onCompleted();
    assertTrue(done.await(30, TimeUnit.SECONDS), "parse timed out");
    return new ParseResult(events, failure.get());
  }

  private static ParseResult parseOk(byte[] bytes, String documentId) throws Exception {
    ParseResult result = parse(bytes, documentId, bytes.length);
    assertEquals(null, result.error(), "parse must succeed: " + result.error());
    assertTrue(result.events().get(0).hasDocumentInfo(), "first event is document info");
    return result;
  }

  // --- tests --------------------------------------------------------------

  @Test
  void docxRoundTrip() throws Exception {
    ParseResult result = parseOk(docxFixture(), "docx-1");
    var info = result.events().get(0).getDocumentInfo();
    assertEquals("docx-1", info.getDocumentId());
    assertEquals(DocumentFormat.DOCUMENT_FORMAT_DOCX, info.getFormat());
    assertEquals("Quarterly Narrative", info.getMetadata().getTitle());
    assertEquals("Archivist", info.getMetadata().getAuthor());
    assertTrue(
        info.getMetadata().getTailList().stream()
            .anyMatch(entry -> entry.getKey().equals("custom:approved")
                && entry.getValues(0).getBoolValue()),
        "custom property must arrive typed");

    List<ParseEvent> paragraphs =
        result.events().stream().filter(ParseEvent::hasParagraph).toList();
    assertEquals("Introduction", paragraphs.get(0).getParagraph().getText());
    assertEquals("Heading1", paragraphs.get(0).getParagraph().getStyle());
    List<ParseEvent> tables = result.events().stream().filter(ParseEvent::hasTable).toList();
    assertEquals(1, tables.size());
    var table = tables.get(0).getTable();
    assertEquals(2, table.getRowsCount());
    assertEquals("Region", table.getRows(0).getCells(0).getText());
    assertEquals("1200", table.getRows(1).getCells(1).getText());
    assertEquals(ParseStatus.State.STATE_OK, result.status().getState());
    assertEquals(2, result.status().getParagraphs());
    assertEquals(1, result.status().getTables());
  }

  @Test
  void xlsxRoundTripKeepsTypes() throws Exception {
    ParseResult result = parseOk(xlsxFixture(), "xlsx-1");
    assertEquals(DocumentFormat.DOCUMENT_FORMAT_XLSX,
        result.events().get(0).getDocumentInfo().getFormat());
    List<ParseEvent> sheets = result.events().stream().filter(ParseEvent::hasSheet).toList();
    assertEquals(2, sheets.size());
    var data = sheets.get(0).getSheet();
    assertEquals("Data", data.getName());
    var valueRow = data.getRows(1);
    SheetCell amount = valueRow.getCells(0);
    assertEquals(42.5, amount.getNumber());
    assertTrue(valueRow.getCells(1).getBoolean());
    SheetCell formula = valueRow.getCells(2);
    assertEquals("A2*2", formula.getFormula());
    assertEquals(85.0, formula.getNumber(), "cached formula result rides the typed value");
    assertEquals("Empty", sheets.get(1).getSheet().getName());
    assertEquals(0, sheets.get(1).getSheet().getRowsCount());
    assertEquals(2, result.status().getSheets());
  }

  @Test
  void pptxRoundTrip() throws Exception {
    ParseResult result = parseOk(pptxFixture(), "pptx-1");
    assertEquals(DocumentFormat.DOCUMENT_FORMAT_PPTX,
        result.events().get(0).getDocumentInfo().getFormat());
    List<ParseEvent> slides = result.events().stream().filter(ParseEvent::hasSlide).toList();
    assertEquals(1, slides.size());
    assertEquals(List.of("The archivist's dream machine"),
        slides.get(0).getSlide().getTextsList());
  }

  @Test
  void legacyXlsAndPptRoundTrip() throws Exception {
    ParseResult xls = parseOk(xlsFixture(), "xls-1");
    assertEquals(DocumentFormat.DOCUMENT_FORMAT_LEGACY_XLS,
        xls.events().get(0).getDocumentInfo().getFormat());
    assertEquals("still works",
        xls.events().stream().filter(ParseEvent::hasSheet).findFirst().orElseThrow()
            .getSheet().getRows(0).getCells(0).getText());

    ParseResult ppt = parseOk(pptFixture(), "ppt-1");
    assertEquals(DocumentFormat.DOCUMENT_FORMAT_LEGACY_PPT,
        ppt.events().get(0).getDocumentInfo().getFormat());
    assertEquals(List.of("legacy slide text"),
        ppt.events().stream().filter(ParseEvent::hasSlide).findFirst().orElseThrow()
            .getSlide().getTextsList());
  }

  @Test
  void chunkedUploadMatchesSingleChunk() throws Exception {
    byte[] bytes = docxFixture();
    ParseResult whole = parse(bytes, "doc", bytes.length);
    ParseResult chunked = parse(bytes, "doc", 1024);
    assertEquals(null, chunked.error());
    assertEquals(whole.events().size(), chunked.events().size(),
        "chunking must not change the event stream");
  }

  @Test
  void garbageBytesAreUnimplemented() throws Exception {
    byte[] noise = new byte[512];
    for (int index = 0; index < noise.length; index++) noise[index] = (byte) (index * 31);
    ParseResult result = parse(noise, "junk", noise.length);
    assertNotNull(result.error());
    assertEquals(Status.Code.UNIMPLEMENTED,
        ((StatusRuntimeException) result.error()).getStatus().getCode());
  }

  @Test
  void oversizeDocumentIsResourceExhausted() throws Exception {
    byte[] big = new byte[(int) CAP_BYTES + 1];
    ParseResult result = parse(big, "big", big.length);
    assertNotNull(result.error());
    assertEquals(Status.Code.RESOURCE_EXHAUSTED,
        ((StatusRuntimeException) result.error()).getStatus().getCode());
  }

  @Test
  void missingCompleteFlagIsInvalid() throws Exception {
    byte[] bytes = docxFixture();
    List<ParseEvent> events = new ArrayList<>();
    AtomicReference<Throwable> failure = new AtomicReference<>();
    CountDownLatch done = new CountDownLatch(1);
    StreamObserver<ParseRequestChunk> requests =
        PoiParseServiceGrpc.newStub(channel).parseDocument(new StreamObserver<>() {
          @Override
          public void onNext(ParseEvent event) {
            events.add(event);
          }

          @Override
          public void onError(Throwable error) {
            failure.set(error);
            done.countDown();
          }

          @Override
          public void onCompleted() {
            done.countDown();
          }
        });
    requests.onNext(ParseRequestChunk.newBuilder().setDocumentId("partial")
        .setData(ByteString.copyFrom(bytes)).build());
    requests.onCompleted();
    assertTrue(done.await(10, TimeUnit.SECONDS));
    assertNotNull(failure.get());
    assertEquals(Status.Code.INVALID_ARGUMENT,
        ((StatusRuntimeException) failure.get()).getStatus().getCode());
    assertTrue(events.isEmpty(), "no events before validation");
  }

  @Test
  void serviceInfoReportsCapabilities() {
    GetServiceInfoResponse info = PoiParseServiceGrpc.newBlockingStub(channel)
        .getServiceInfo(GetServiceInfoRequest.getDefaultInstance());
    assertEquals(PoiParseServiceImpl.SERVICE_VERSION, info.getServiceVersion());
    assertFalse(info.getPoiVersion().isEmpty());
    assertEquals(6, info.getSupportedFormatsCount());
    assertEquals(CAP_BYTES, info.getMaxDocumentBytes());
    assertEquals(4, info.getMaxConcurrentParses());
    assertEquals("POI", info.getUi().getTitle());
    assertEquals("/ui/grpoic", info.getUi().getPath());
    assertFalse(info.getUi().getDescription().isEmpty());
  }

  @Test
  void concurrentParsesComplete() throws Exception {
    byte[] docx = docxFixture();
    byte[] xlsx = xlsxFixture();
    List<Thread> threads = new ArrayList<>();
    AtomicReference<Throwable> firstFailure = new AtomicReference<>();
    for (int index = 0; index < 8; index++) {
      byte[] bytes = index % 2 == 0 ? docx : xlsx;
      String id = "concurrent-" + index;
      threads.add(Thread.ofVirtual().start(() -> {
        try {
          ParseResult result = parse(bytes, id, 2048);
          if (result.error() != null) firstFailure.compareAndSet(null, result.error());
        } catch (Throwable error) {
          firstFailure.compareAndSet(null, error);
        }
      }));
    }
    for (Thread thread : threads) thread.join(TimeUnit.SECONDS.toMillis(30));
    assertEquals(null, firstFailure.get(), "all concurrent parses must succeed");
  }
}
