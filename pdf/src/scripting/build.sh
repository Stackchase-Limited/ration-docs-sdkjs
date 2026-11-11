#!/bin/sh
# Rebuilds ../engine/pdfscripting.js from sandbox.js and the pinned packages in package.json.
set -e
cd "$(dirname "$0")"
npm ci --ignore-scripts --no-audit --no-fund
OUT=../engine/pdfscripting.js
{
	printf '/* Ration Docs PDF scripting sandbox, built from pdf/src/scripting/sandbox.js by build.sh.\n'
	printf ' * Contains QuickJS (MIT) and quickjs-emscripten (MIT); licence text follows. */\n/*\n'
	sed 's/\*\//* \//g' node_modules/@jitl/quickjs-singlefile-browser-release-sync/LICENSE
	printf '*/\n'
	node node_modules/esbuild/bin/esbuild sandbox.js --bundle --format=iife --minify --target=chrome109 --legal-comments=none
} > "$OUT"
rm -rf node_modules
