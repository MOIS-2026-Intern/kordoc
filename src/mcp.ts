/**
 * kordoc MCP 서버 — Claude/Cursor에서 문서 파싱 도구로 사용
 *
 * 사용법:
 *   node dist/mcp.js              (stdio 모드)
 *   node dist/mcp.js --http 3000  (HTTP 모드)
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { z } from "zod"
import { readFileSync } from "fs"
import { parse, detectFormat } from "./index.js"

const server = new McpServer({
  name: "kordoc",
  version: "0.1.0",
})

// ─── 도구: parse_document ────────────────────────────

server.tool(
  "parse_document",
  "한국 문서 파일(HWP, HWPX, PDF)을 마크다운으로 변환합니다. 파일 경로를 입력하면 포맷을 자동 감지하여 텍스트를 추출합니다.",
  {
    file_path: z.string().describe("파싱할 문서 파일의 절대 경로 (HWP, HWPX, PDF)"),
  },
  async ({ file_path }) => {
    try {
      const buffer = readFileSync(file_path)
      const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)
      const format = detectFormat(arrayBuffer)

      if (format === "unknown") {
        return {
          content: [{ type: "text", text: `지원하지 않는 파일 형식입니다: ${file_path}` }],
          isError: true,
        }
      }

      const result = await parse(arrayBuffer)

      if (!result.success) {
        return {
          content: [{ type: "text", text: `파싱 실패 (${result.fileType}): ${result.error}` }],
          isError: true,
        }
      }

      const meta = [
        `포맷: ${result.fileType.toUpperCase()}`,
        result.pageCount ? `페이지: ${result.pageCount}` : null,
        result.isImageBased ? "이미지 기반 PDF (텍스트 추출 불가)" : null,
      ].filter(Boolean).join(" | ")

      return {
        content: [{ type: "text", text: `[${meta}]\n\n${result.markdown}` }],
      }
    } catch (err) {
      return {
        content: [{ type: "text", text: `오류: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      }
    }
  }
)

// ─── 도구: detect_format ─────────────────────────────

server.tool(
  "detect_format",
  "파일의 포맷을 매직 바이트로 감지합니다 (hwpx, hwp, pdf, unknown).",
  {
    file_path: z.string().describe("감지할 파일의 절대 경로"),
  },
  async ({ file_path }) => {
    try {
      const buffer = readFileSync(file_path, { flag: "r" })
      const header = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + Math.min(buffer.byteLength, 16))
      const format = detectFormat(header)
      return {
        content: [{ type: "text", text: `${file_path}: ${format}` }],
      }
    } catch (err) {
      return {
        content: [{ type: "text", text: `오류: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      }
    }
  }
)

// ─── 서버 시작 ───────────────────────────────────────

async function main() {
  const transport = new StdioServerTransport()
  await server.connect(transport)
}

main().catch(console.error)
