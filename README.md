# beamer-to-pptx

Command line tool to convert Beamer-generated PDFs to Powerpoint files, using SVGs for lossless conversion and including speaker notes.

Requirements:

  - Poppler PDF tools, with `pdfinfo`, `pdftocairo` and `pdftotext` available in your PATH.
  - Node/Javascript runtime and `npm` (or similar) to install the Javscript dependencies.
  
Usage:

  1. Make sure you have `npm` (with `npx`) and Poppler (with `pdfinfo`, `pdftocairo` and `pdftotext`) installed and working. 
  2. Run `npx beamer-to-pptx` in your shell to run a temporary installation of the tool, which should output the usage instructions:
```
  _                                         _                          _                   
 | |__   ___  __ _ _ __ ___   ___ _ __     | |_ ___        _ __  _ __ | |___  __           
 | '_ \ / _ \/ _` | '_ ` _ \ / _ \ '__|____| __/ _ \ _____| '_ \| '_ \| __\ \/ /  
 | |_) |  __/ (_| | | | | | |  __/ | |_____| || (_) |_____| |_) | |_) | |_ >  <            
 |_.__/ \___|\__,_|_| |_| |_|\___|_|        \__\___/      | .__/| .__/ \__/_/\_\   
                                                          |_|   |_|                        
 
 Convert Beamer slides to Powerpoint presentations losslessly*, including your 
 speaker notes! 
 
 Usage: 
   1. Compile your Beamer slides with: 
      \setbeameroption{show notes on second screen=right} 
      \setbeamertemplate{note page}[plain] 
   2. Convert the PDF into a Powerpoint presentation: 
      beamer-to-pptx <presentation.pdf> 
 
 This utility uses poppler's pdftocairo to render each slide to an SVG (with 
 fonts converted to glyphs to avoid breakage). It then runs pdftotext to get 
 the textual content of the speaker notes. These are combined into a Powerpoint 
 presentation with each slide consisting of a fullscreen SVG and speaker notes. 
 
 The resulting Powerpoint presentation is saved as <presentation>.pptx with 
 <presentation> taken from the input file name <presentation.pdf>. The output is 
 intended to be read-only (so you can use speaker prompts at venues that only 
 support Powerpoint). They aren't editable in any meaningful way. 
 
 Note that your presentation must be in 4x3, 16x10 or 16x9 aspect ratio due to 
 some bugs I haven't been able to fix. 
 
 * If `pdftocairo -svg` is lossless for your PDF, then this tool should be as 
   well. The `pdftotext` transformation of speaker notes is likely to be lossy. 
```

In short, make sure Poppler and npm/npx are installed then run `npx beamer-to-pptx <path-to-presentation.pdf>`. You could alternatively install the package as you would any other npm package.
