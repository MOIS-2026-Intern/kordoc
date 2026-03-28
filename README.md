# kordoc

### 모두 파싱해버리겠다.

> *"HWP든 HWPX든 PDF든 — 대한민국 문서라면 남김없이 파싱해버립니다."*

Built by a Korean civil servant who spent 7 years in the deepest circle of document hell. One day he snapped, and kordoc was born.

Korean document formats — parsed, converted, delivered as clean Markdown. No COM automation, no Windows dependency, no tears.

---

## Why kordoc?

South Korea runs on HWP. The rest of the world has never heard of it. Government offices produce thousands of `.hwp` files daily, and extracting text from them has always been a nightmare — COM automation that only works on Windows, proprietary formats with zero documentation, and tables that break every parser.

**kordoc** was forged in this document hell. Its parsers have been battle-tested across 5 real Korean government projects, processing everything from school curriculum plans to facility inspection reports. If a Korean public servant wrote it, kordoc can parse it.

| Format | Engine | Status |
|--------|--------|--------|
| **HWPX** (한컴 2020+) | ZIP + XML DOM walk | Stable |
| **HWP 5.x** (한컴 레거시) | OLE2 binary + record parsing | Stable |
| **PDF** | pdfjs-dist text extraction | Stable |

### What makes it different

- **2-pass table builder** — Correct `colSpan`/`rowSpan` handling via grid algorithm. No more broken table layouts.
- **Broken ZIP recovery** — Corrupted HWPX? We scan raw Local File Headers and still extract text.
- **OPF manifest resolution** — Multi-section HWPX documents parsed in correct spine order.
- **21 HWP5 control characters** — Full UTF-16LE decoding with extended/inline object skip.
- **Image-based PDF detection** — Warns you when a scanned PDF can't be text-extracted.

---

## Quick Start

### As a library

```bash
npm install kordoc
```

```typescript
import { parse } from "kordoc"
import { readFileSync } from "fs"

const buffer = readFileSync("document.hwpx")
const result = await parse(buffer.buffer)

if (result.success) {
  console.log(result.markdown)
  // → Clean markdown with tables, headings, and structure preserved
}
```

### Format-specific parsing

```typescript
import { parseHwpx, parseHwp, parsePdf } from "kordoc"

// HWPX (modern Hancom format)
const hwpxResult = await parseHwpx(buffer)

// HWP 5.x (legacy binary format)
const hwpResult = await parseHwp(buffer)

// PDF (text-based)
const pdfResult = await parsePdf(buffer)
```

### Format detection

```typescript
import { detectFormat, isHwpxFile, isOldHwpFile, isPdfFile } from "kordoc"

const format = detectFormat(buffer) // → "hwpx" | "hwp" | "pdf" | "unknown"
```

### As a CLI

```bash
npx kordoc document.hwpx                    # stdout
npx kordoc document.hwp -o output.md        # save to file
npx kordoc *.pdf -d ./converted/            # batch convert
npx kordoc report.hwpx --format json        # JSON with metadata
```

### As an MCP server (Claude / Cursor / Windsurf)

kordoc includes a built-in MCP server. Add it to your Claude Desktop config:

```json
{
  "mcpServers": {
    "kordoc": {
      "command": "npx",
      "args": ["-y", "kordoc-mcp"]
    }
  }
}
```

This exposes two tools:
- **`parse_document`** — Parse a HWP/HWPX/PDF file to Markdown
- **`detect_format`** — Detect file format via magic bytes

---

## API Reference

### `parse(buffer: ArrayBuffer): Promise<ParseResult>`

Auto-detects format via magic bytes and parses to Markdown.

### `ParseResult`

```typescript
interface ParseResult {
  success: boolean
  markdown?: string          // Extracted markdown text
  fileType: "hwpx" | "hwp" | "pdf" | "unknown"
  isImageBased?: boolean     // true if scanned PDF (no text extractable)
  pageCount?: number         // PDF page count
  error?: string             // Error message on failure
}
```

### Low-level exports

```typescript
// Table builder (2-pass colSpan/rowSpan algorithm)
import { buildTable, blocksToMarkdown } from "kordoc"

// Type definitions
import type { IRBlock, IRTable, IRCell, CellContext } from "kordoc"
```

---

## Supported Formats

### HWPX (한컴오피스 2020+)

ZIP-based XML format. kordoc reads the OPF manifest (`content.hpf`) for correct section ordering, walks the XML DOM for paragraphs and tables, and handles:
- Multi-section documents
- Nested tables (table inside a table cell)
- `colSpan` / `rowSpan` merged cells
- Corrupted ZIP archives (Local File Header fallback)

### HWP 5.x (한컴오피스 레거시)

OLE2 Compound Binary format. kordoc parses the CFB container, decompresses section streams (zlib), reads HWP record structures, and extracts UTF-16LE text with full control character handling:
- 21 control character types (line breaks, tabs, hyphens, NBSP, extended objects)
- Encrypted/DRM file detection (fails fast with clear error)
- Table extraction with grid-based cell arrangement

### PDF

Server-side text extraction via pdfjs-dist:
- Y-coordinate based line grouping
- Gap-based cell/table detection
- Image-based PDF detection (< 10 chars/page average)
- Korean text line joining (조사/접속사 awareness)

---

## Requirements

- **Node.js** >= 18
- **pdfjs-dist** — Required only for PDF parsing. HWP/HWPX work without it.

---

## Credits

Built by a Korean civil servant who spent years drowning in HWP files. Production-tested across 5 government technology projects — school curriculum plans, facility inspection reports, legal documents, and municipal newsletters. The parsers in this package have processed thousands of real Korean government documents without breaking a sweat.

---

## License

MIT

---

<br>

# kordoc (한국어)

### 모두 파싱해버리겠다.

> *대한민국에서 둘째가라면 서러울 문서지옥. 거기서 7년 버틴 공무원이 만들었습니다.*

HWP, HWPX, PDF — 관공서에서 쏟아지는 모든 문서 포맷을 마크다운으로 변환하는 Node.js 라이브러리입니다. 학교 교육과정, 사전기획 보고서, 검토의견서, 소식지 원고... 뭐든 넣으면 파싱합니다.

### 특징

- **한컴오피스 불필요** — COM 자동화 없이 바이너리 직접 파싱. Linux, Mac에서도 동작
- **손상 파일 복구** — ZIP Central Directory가 깨진 HWPX도 Local File Header 스캔으로 복구
- **병합 셀 완벽 처리** — 2-pass 그리드 알고리즘으로 colSpan/rowSpan 정확히 렌더링
- **HWP5 바이너리 직접 파싱** — OLE2 컨테이너 → 레코드 스트림 → UTF-16LE 텍스트 추출
- **이미지 PDF 감지** — 스캔된 PDF는 텍스트 추출 불가를 사전에 알려줌
- **실전 검증 완료** — 5개 공공 프로젝트, 수천 건의 실제 관공서 문서에서 테스트됨

### 설치

```bash
npm install kordoc
```

### 사용법

```typescript
import { parse } from "kordoc"
import { readFileSync } from "fs"

const buffer = readFileSync("사업계획서.hwpx")
const result = await parse(buffer.buffer)

if (result.success) {
  console.log(result.markdown)
}
```

### CLI

```bash
npx kordoc 사업계획서.hwpx                     # 터미널 출력
npx kordoc 보고서.hwp -o 보고서.md              # 파일 저장
npx kordoc *.pdf -d ./변환결과/                 # 일괄 변환
```

### MCP 서버 (Claude / Cursor / Windsurf)

Claude Desktop이나 Cursor에서 문서 파싱 도구로 바로 사용 가능합니다:

```json
{
  "mcpServers": {
    "kordoc": {
      "command": "npx",
      "args": ["-y", "kordoc-mcp"]
    }
  }
}
```
