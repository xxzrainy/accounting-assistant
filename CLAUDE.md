# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

An accounting assistant built with Google Apps Script (GAS), managed via [clasp](https://github.com/nicholaschiang/clasp).

## Development

- GAS source files live in `./src/` (configured as `rootDir` in `.clasp.json`)
- Use `clasp push` to deploy code to Google Apps Script
- Use `clasp pull` to pull latest from the remote script
- Use `clasp open` to open the script in the Apps Script editor

## Key Notes

- This project uses clasp for local GAS development; all script files should be placed under `src/`
- GAS files use `.gs` extension (or `.ts` if using TypeScript with clasp)
