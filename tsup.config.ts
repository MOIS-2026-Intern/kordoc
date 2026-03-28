import { defineConfig } from "tsup"

export default defineConfig([
  {
    entry: ["src/index.ts"],
    format: ["esm", "cjs"],
    dts: true,
    splitting: false,
    sourcemap: true,
    clean: true,
    external: ["pdfjs-dist"],
    noExternal: ["cfb"],
  },
  {
    entry: ["src/cli.ts", "src/mcp.ts"],
    format: ["esm"],
    banner: { js: "#!/usr/bin/env node" },
    external: ["pdfjs-dist"],
    noExternal: ["cfb"],
  },
])
