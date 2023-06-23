# beamer-to-pptx

Command line tool to convert Beamer-generated PDFs to Powerpoint files, using SVGs for lossless conversion and including speaker notes. It assumes the PDFs it is given are Beamer presentations compiled to a PDF with the following preamble:

``` latex
\setbeameroption{show notes on second screen=right}
\setbeamertemplate{note page}[plain]
```

Requirements:

  - Poppler PDF tools, with `pdfinfo`, `pdftocairo` and `pdftotext` available in your PATH.
  - Node/Javascript runtime and `npm` (or similar) to install the Javscript dependencies.

Usage:

  - Assuming you have the pre-requisites above installed, run `npx beamer-to-pptx` to run a temporary installation of the tool and output the usage instructions or `npx beamer-to-pptx <path-to-beamer-slides.pdf>` to do the conversion directly.
  - Alternatively, you could install the tool like any other npm package, or clone the repo and run the script directly.
