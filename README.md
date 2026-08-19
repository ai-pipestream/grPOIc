# grPOIc

A gRPC server that wraps [Apache POI](https://poi.apache.org/). Clients stream
office document bytes in and receive typed structure events back: metadata,
paragraphs, tables, sheets, slides, embedded objects, and a final status.

The server exists for two reasons:

1. POI is a library, not a service. Putting it behind gRPC gives non-JVM
   clients access to office parsing over a stable wire contract, and isolates
   parser crashes and memory use in a separate process.
2. Document bytes should not touch disk. The whole parse happens in memory:
   no temp files, no subprocesses. The container runs with a read-only root
   filesystem.

grPOIc extracts content and metadata. It does not render or convert
documents.

## API

`ai.pipestream.poi.v1.PoiParseService` (see `grpoic-api/src/main/proto`):

- `ParseDocument(stream ParseRequestChunk) returns (stream ParseEvent)`. The
  client streams the document as chunks, marking the last one `complete`. The
  server streams back `DocumentInfo` (detected format plus typed metadata),
  then content blocks in document order (`Paragraph`, `Table`, `Sheet`,
  `Slide`, `EmbeddedObject`), then one final `ParseStatus`.
- `GetServiceInfo`: versions, supported formats, and operational limits, for
  orchestrators and tool facades that need capability discovery. Also carries
  a `UiInfo` block advertising this service's tab to the shared demo shell.

Formats: DOCX, XLSX, PPTX and the OLE2 legacy trio DOC, XLS, PPT. The format
is detected from the bytes; the advisory content type is never trusted.
Spreadsheet cells keep their storage types (string, double, boolean, date).
Formula cells carry the formula source plus the cached result; formulas are
never evaluated. Metadata is typed and lossless: well-known core properties as
first-class fields, everything else in a tagged tail, nothing guessed from
string shapes.

Errors are gRPC status codes: `INVALID_ARGUMENT` (no bytes, missing complete
flag, unreadable claimed format), `RESOURCE_EXHAUSTED` (over the byte cap),
`UNIMPLEMENTED` (not an office format), `INTERNAL` (parser fault). Standard
gRPC health checking and reflection are registered.

## Concurrency model

POI documents are single-threaded; distinct documents on distinct threads is
the supported pattern. Each parse runs on its own virtual thread, with a
semaphore bounding concurrent parses. The bound protects heap (POI holds full
document models in memory), not just CPU.

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

Tests author their fixtures with POI itself, in memory. No binary files are
committed, and every assertion is against content the test placed.
