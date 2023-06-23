#!/usr/bin/env node

const os = require("os");
const pptxgen = require("pptxgenjs");
const { readFileSync } = require("node:fs");
const { hideBin } = require("yargs/helpers");
const { execSync } = require("child_process");

function get_pdf_info(pdf_path) {
    let pdfinfo = execSync("pdfinfo " + pdf_path).toString();

    let page_count_search = pdfinfo.match(/Pages:\s+(\d+)/);
    if (page_count_search == null) {return null};
    if (page_count_search.length != 2) {return null};

    let page_count = parseInt(page_count_search[1]);

    let dimensions_search = pdfinfo.match(/Page size:\s+([\d\.]+)\s*x\s*([\d\.]+)\s+pts/);
    if (dimensions_search == null) {return null};
    if (dimensions_search.length != 3) {return null};

    let width_pts = parseFloat(dimensions_search[1]);
    let height_pts = parseFloat(dimensions_search[2]);

    let title_search = pdfinfo.match(/Title:\s+(.+)\n/);
    let title = '';
    if (title_search != null && title_search.length == 2) {
        title = title_search[1];
    };

    let author_search = pdfinfo.match(/Author:\s+(.+)\n/);
    let author = '';
    if (author_search != null && author_search.length == 2) {
        author = author_search[1];
    };

    return {
        "title": title,
        "author": author,
        "page_count": page_count,
        "page_width": width_pts / 2.0,
        "page_height": height_pts
    };
}

function guess_aspect_ratio(width, height) {
    width = Number(width);
    height = Number(height);

    let ratio = width/height;

    let fourthree = 4.0/3.0;
    let sixteennine = 16.0/9.0;
    let sixteenten = 16.0/10.0;

    let distance_from_fourthree = Math.abs(fourthree - ratio);
    let distance_from_sixteennine = Math.abs(sixteennine - ratio);
    let distance_from_sixteenten = Math.abs(sixteenten - ratio);

    if (distance_from_fourthree < distance_from_sixteennine) {
        if (distance_from_fourthree < distance_from_sixteennine) {
            return "4x3";
        } else {
            return "16x10";
        }
    } else {
        if (distance_from_sixteennine < distance_from_sixteenten) {
            return "16x9";
        } else {
            return "16x10";
        }
    }
}

let args = hideBin(process.argv);
if (args.length != 1) {
    console.log("\
  _                                         _                          _                   \n \
| |__   ___  __ _ _ __ ___   ___ _ __     | |_ ___        _ __  _ __ | |___  __           \n \
| '_ \\ / _ \\/ _` | '_ ` _ \\ / _ \\ '__|____| __/ _ \\ _____| '_ \\| '_ \\| __\\ \\/ /  \n \
| |_) |  __/ (_| | | | | | |  __/ | |_____| || (_) |_____| |_) | |_) | |_ >  <            \n \
|_.__/ \\___|\\__,_|_| |_| |_|\\___|_|        \\__\\___/      | .__/| .__/ \\__/_/\\_\\   \n \
                                                         |_|   |_|                        \n \
\n \
Convert Beamer slides to Powerpoint presentations losslessly*, including your \n \
speaker notes! \n \
\n \
Usage: \n \
  1. Compile your Beamer slides with: \n \
     \\setbeameroption{show notes on second screen=right} \n \
     \\setbeamertemplate{note page}[plain] \n \
  2. Convert the PDF into a Powerpoint presentation: \n \
     beamer-to-pptx <presentation.pdf> \n \
\n \
This utility uses poppler's pdftocairo to render each slide to an SVG (with \n \
fonts converted to glyphs to avoid breakage). It then runs pdftotext to get \n \
the textual content of the speaker notes. These are combined into a Powerpoint \n \
presentation with each slide consisting of a fullscreen SVG and speaker notes. \n \
\n \
The resulting Powerpoint presentation is saved as <presentation>.pptx with \n \
<presentation> taken from the input file name <presentation.pdf>. The output is \n \
intended to be read-only (so you can use speaker prompts at venues that only \n \
support Powerpoint). They aren't editable in any meaningful way. \n \
\n \
Note that your presentation must be in 4x3, 16x10 or 16x9 aspect ratio due to \n \
some bugs I haven't been able to fix. \n \
\n \
* If `pdftocairo -svg` is lossless for your PDF, then this tool should be as \n \
  well. The `pdftotext` transformation of speaker notes is likely to be lossy. \n \
    ")
    process.exit(1);
}

let pdf_path = args[0];
let pptx_path;
if (pdf_path.endsWith(".pdf")) { pptx_path = pdf_path.substring(0, pdf_path.length - 4) }
else { pptx_path = pdf_path; }
pptx_path += ".pptx";

let pdf_info = get_pdf_info(pdf_path);
let page_count = pdf_info["page_count"];
let width = pdf_info["page_width"];
let height = pdf_info["page_height"];

// Convert to SVG's
let slides_cropbox = "-nocenter -x 0 -y 0 -W " + width.toFixed() + " -H " + height.toFixed();

let svg_dir = os.tmpdir();
let svg_paths = [];
for (let page = 1; page < page_count+1; page++) {
    svg_paths[page] = svg_dir + page.toString() + ".svg";
    let cmd = "pdftocairo -svg " + slides_cropbox + " -f " + page.toString() + " -l " + page.toString() + " " + pdf_path + " " + svg_paths[page];
    console.log(cmd);
    execSync(cmd);
}

// Extract text from notes
let notes_cropbox = " -x " + width.toFixed() + " -y 0 -W " + width.toFixed() + " -H " + height.toFixed();

let txt_dir = os.tmpdir();
let txt_paths = [];
for (let page = 1; page < page_count+1; page++) {
    txt_paths[page] = txt_dir + page.toString() + ".txt";
    let cmd = "pdftotext -nodiag -nopgbrk " + notes_cropbox + " -f " + page.toString() + " -l " + page.toString() + " " + pdf_path + " " + txt_paths[page];
    console.log(cmd);
    execSync(cmd);
}

// Construct presentation
let pres = new pptxgen();

pres.title = pdf_info["title"];
pres.author = pdf_info["author"];
pres.subject = "";
pres.company = "";

pres.layout = "LAYOUT_" + guess_aspect_ratio(width, height);

// Build it slide by slide
for (let page = 1; page < page_count+1; page++) {
    let slide = pres.addSlide();

    slide.addImage({
        path: svg_paths[page],
        x: 0,
        y: 0,
        w: "200%",  // HACK: when we crop with pdftocairo the page size stays at double width (with a bunch of whitespace where the cropped-out stuff is).
        h: "100%"
    });

    let notes = readFileSync(txt_paths[page], {encoding: "utf8"});
    slide.addNotes(notes);
}

// Save!
pres.writeFile({ fileName: pptx_path });
