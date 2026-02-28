#!/usr/bin/env bash
# SPDX-License-Identifier: MIT
# Build CodeMirror 6 IIFE bundle for OpenSpace
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
OUT="$PROJECT_ROOT/web/static/lib/codemirror.bundle.js"

TMPDIR="$(mktemp -d)"
trap 'rm -rf "$TMPDIR"' EXIT

cd "$TMPDIR"
npm init -y >/dev/null 2>&1

npm install --save esbuild \
  @codemirror/state @codemirror/view @codemirror/commands \
  @codemirror/language @codemirror/search @codemirror/autocomplete \
  @codemirror/lang-javascript @codemirror/lang-python @codemirror/lang-html \
  @codemirror/lang-css @codemirror/lang-markdown @codemirror/lang-xml \
  @codemirror/lang-sql @codemirror/lang-json \
  @codemirror/theme-one-dark @lezer/highlight 2>&1 | tail -5

cat > entry.js << 'ENTRY'
// CodeMirror 6 IIFE bundle entry
import {EditorState, Compartment, StateEffect, StateField} from "@codemirror/state";
import {EditorView, keymap, lineNumbers, highlightActiveLine, highlightActiveLineGutter,
        drawSelection, rectangularSelection, crosshairCursor, dropCursor,
        highlightSpecialChars, placeholder} from "@codemirror/view";
import {defaultKeymap, history, historyKeymap, indentWithTab} from "@codemirror/commands";
import {syntaxHighlighting, defaultHighlightStyle, indentOnInput, bracketMatching,
        foldGutter, foldKeymap, HighlightStyle, indentUnit} from "@codemirror/language";
import {searchKeymap, highlightSelectionMatches, search, openSearchPanel} from "@codemirror/search";
import {autocompletion, completionKeymap, closeBrackets, closeBracketsKeymap} from "@codemirror/autocomplete";
import {javascript} from "@codemirror/lang-javascript";
import {python} from "@codemirror/lang-python";
import {html} from "@codemirror/lang-html";
import {css} from "@codemirror/lang-css";
import {markdown} from "@codemirror/lang-markdown";
import {xml} from "@codemirror/lang-xml";
import {sql} from "@codemirror/lang-sql";
import {json} from "@codemirror/lang-json";
import {oneDark} from "@codemirror/theme-one-dark";
import {tags} from "@lezer/highlight";

globalThis.CM = {
  EditorState, EditorView, Compartment, StateEffect, StateField,
  keymap, lineNumbers, highlightActiveLine, highlightActiveLineGutter,
  drawSelection, rectangularSelection, crosshairCursor, dropCursor,
  highlightSpecialChars, placeholder,
  defaultKeymap, history, historyKeymap, indentWithTab,
  syntaxHighlighting, defaultHighlightStyle, indentOnInput, bracketMatching,
  foldGutter, foldKeymap, HighlightStyle, indentUnit,
  searchKeymap, highlightSelectionMatches, search, openSearchPanel,
  autocompletion, completionKeymap, closeBrackets, closeBracketsKeymap,
  javascript, python, html, css, markdown, xml, sql, json,
  oneDark, tags,
};
ENTRY

npx esbuild entry.js --bundle --format=iife --minify --outfile="$OUT" 2>&1

echo ""
echo "Built: $OUT"
ls -lh "$OUT"
