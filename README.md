# grPOIc

A diskless gRPC streaming server over Apache POI: office document bytes in,
typed structure events out. No Tika, no temp files, no `exec()` — document
bytes live only in memory for the duration of the parse.

grPOIc is the structured fast path of the pipestream office pipeline. Its
sibling, the office-bridge (pooled LibreOffice), owns rendering fidelity and
office-to-PDF conversion; [gRParse](https://github.com/ai-pipestream/gRParse)
consumes those PDFs for OCR, layout, and page images. grPOIc answers the
question "what does this document say and contain" without rendering anything.

## API

`ai.pipestream.poi.v1.PoiParseService` (see `grpoic-api/src/main/proto`):

- `ParseDocument(stream ParseRequestChunk) → stream ParseEvent` — the client
  streams the document as chunks (last one marked `complete`); the server
  streams back `DocumentInfo` (detected format + typed metadata), then content
  blocks in document order (`Paragraph`, `Table`, `Sheet`, `Slide`,
  `EmbeddedObject`), then one final `ParseStatus`.
- `GetServiceInfo` — capability discovery: versions, supported formats, and
  operational limits, intended for orchestrators and LLM tool facades.

Formats: DOCX, XLSX, PPTX and the OLE2 legacy trio DOC, XLS, PPT. The format
is detected from the bytes; the advisory content type is never trusted.
Spreadsheet cells keep their storage types (string, double, boolean, date) and
formula cells carry the formula source plus the cached result — formulas are
never evaluated. Metadata is typed and lossless: well-known core properties as
first-class fields, everything else in a tagged tail, nothing guessed from
string shapes.

Errors are gRPC status codes: `INVALID_ARGUMENT` (no bytes, missing complete
flag, unreadable claimed format), `RESOURCE_EXHAUSTED` (over the byte cap),
`UNIMPLEMENTED` (not an office format), `INTERNAL` (parser fault). Standard
gRPC health checking and reflection are registered.

## Concurrency model

POI documents are single-threaded; distinct documents on distinct threads is
the supported pattern. Each parse runs on its own virtual thread with a
semaphore bounding concurrent parses — the bound protects heap (POI is
memory-hungry), not just CPU.

## Configuration

| Variable | Default | Meaning |
|---|---|---|
| `GRPOIC_PORT` | `50052` | Listen port |
| `GRPOIC_MAX_DOCUMENT_MIB` | `70` | Per-document byte cap (`RESOURCE_EXHAUSTED` above it) |
| `GRPOIC_MAX_CONCURRENT_PARSES` | CPU cores | Parses in flight before queueing |
| `GRPOIC_METRICS_INTERVAL_SECONDS` | `60` | Metrics line interval, `0` disables |

## Build and run

```bash
./gradlew build          # compiles and runs the test suite
./gradlew :grpoic-service:run

docker build -t grpoic .
docker run --rm --read-only -p 50052:50052 grpoic
```

The image build runs the full test suite; `--read-only` works because the
server never writes.

Tests author their fixtures with POI itself, in memory — no binary files are
committed, and every assertion is against content the test placed.
