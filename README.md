# docx

Haskell representation, and writer, of simple Office Open XML Wordpressing
documents

[ECMA-376-1:2016](https://www.ecma-international.org/publications-and-standards/standards/ecma-376)
Office Open XML File Formats (5th Edition) December 2016 (Part 1 — Fundamentals
and Markup Language Reference, October 2016) specifies Office Open XML documents
in the Wordprocessing category.

The representation of such documents is incomplete. The following parts are
represented: the Main Document, Style Definitions, Headers and Footers.

## Other packages

[pandoc](https://hackage.haskell.org/package/pandoc) can parse a `.docx` file to
a `Pandoc` value and write a `Pandoc` value as a `.docx` file.
