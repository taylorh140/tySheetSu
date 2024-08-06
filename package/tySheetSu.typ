#let ss = plugin("tySheetSu.wasm")

#let parse-sheet(xlsxBytes, sheetname) = {
  let xlsx = (json: json.decode(ss.excel_to_json(xlsxBytes, bytes(sheetname))))
  return xlsx
}

#let parse-sheet-data(xlsxBytes, sheetname) = {
  return json.decode(ss.excel_to_json_values(xlsxBytes, bytes(sheetname))).cells
}

#let parse-range(rng, use-merge: true) = {
  // Split the rng into start and end parts
  let rng = upper(rng)
  let parts = rng.split(":")
  let (L1, R1) = parts.first().match(regex("([a-zA-Z]+)([0-9]+)")).captures
  let (L2, R2) = parts.last( ).match(regex("([a-zA-Z]+)([0-9]+)")).captures

  let C1 = range(L1.len())
    .map(x => (
        calc.pow(26, x) * (L1.rev().at(x).to-unicode() - "A".to-unicode() + 1)
      ) )
    .sum()
  let C2 = range(L2.len())
    .map(x => (
        calc.pow(26, x) * (L2.rev().at(x).to-unicode() - "A".to-unicode() + 1)
      ) )
    .sum()
  return (C1, int(R1), C2, int(R2))
}

#let BorderStrokes = (
  thin: 0.5pt + black,
  medium: 1.5pt + black,
  "none": none,
)


#let rangeFn = range;
#let render-table(
  xlsx,
  range: none,
  scale : 1,
  use-merge: true,
  use-color: true,
  use-boarders: true,
  default-padding: 2pt,
  default-boarders: 1pt,
) = {
  let rng = range
  let range = rangeFn

  // Get sub range of table
  let tabledata = none
  if rng == none {
    tabledata = xlsx.json.cells
  } else {
    let (c1, r1, c2, r2) = parse-range(rng)
    tabledata = xlsx.json.cells.filter(cell => (
      cell.x >= c1 and cell.x <= c2 and cell.y >= r1 and cell.y <= r2
    ))
  }

  let align-horizontal-transforms = (
    general : left,
    left   : left,
    center : center,
    right : right,
  )

  let align-vertical-transforms = (
    top: top,
    center: horizon,
    bottom: bottom,
  )
  

  // Use excel scaling (this is hacky)
  let weirdExcelScaleX = 5 / 8.43 * 1em
  let weirdExcelScaleY = 5 / 38.4 * 1em
  let weridExcelRowHeightDefault = 14.4 * scale
  let weirdExcelColWidthDefault = 8.75 * scale

  //Find column row dimensions
  let min-Col = calc.min(..tabledata.map(x => x.x))
  let min-Row = calc.min(..tabledata.map(x => x.y))
  let columns = calc.max(..tabledata.map(x => x.x)) - min-Col + 1
  let rows = calc.max(..tabledata.map(x => x.y)) - min-Row + 1

  //Set column widths based on input using werid number for default.
  let colwidths = range(1,columns + 1)
    .map(x => xlsx.json.columns.filter(row => row.x == x).at(
        0,
        default: (height: weirdExcelColWidthDefault),
      ))
    .map(x => if ("width" in x and x.width == 0 and "hidden" in x and x.hidden == false) {
    weirdExcelColWidthDefault * weirdExcelScaleX
  } else if "hidden" in x and x.hidden==true { 0pt}
  else if "width" in x {
    x.width * weirdExcelScaleX
  } else {
    weridExcelRowHeightDefault * weirdExcelScaleX
  })

  //Set row heights based on excel input using odd number for default
  let rowHeights = range(1,rows + 1)
    .map(y => xlsx.json.rows.filter(row => row.y == y).at(
        0,
        default: (height: weridExcelRowHeightDefault),
      ))
    .map(x => if (x.height == 0 and x.hidden == false) {
    weridExcelRowHeightDefault * weirdExcelScaleY
  } else if "hidden" in x and x.hidden==true { 0pt}
  else if "height" in x {
    x.height * weirdExcelScaleY
  } else {
    weridExcelRowHeightDefault * weirdExcelScaleY
  })

  //Draw the table using the information
  table(columns: colwidths,
    rows: rowHeights,
    stroke: default-boarders,
    ..tabledata.map(cel => table.cell(
      x: cel.x - min-Col,
      y: cel.y - min-Row,
      inset:default-padding,
      stroke: if ("boarders" in cel) and use-boarders {
        (
          left: BorderStrokes.at(cel.boarders.left),
          right: BorderStrokes.at(cel.boarders.right),
          top: BorderStrokes.at(cel.boarders.top),
          bottom: BorderStrokes.at(cel.boarders.bottom),
        )
      } else {
        if not use-boarders {
          default-boarders
        } else {
          none
        }
      },
      colspan: if ("w" in cel and use-merge) {
        cel.w
      } else {
        1
      },
      rowspan: if ("h" in cel and use-merge) {
        cel.h
      } else {
        1
      },
      fill: if (
        ("fill_color" in cel) and (cel.fill_color != "") and use-color
      ) {
        let tmp = cel.fill_color
        while(tmp.len() < 6){
          tmp = "0" + tmp
        }
        color.rgb("#" + tmp.slice(-6))
      } else {
        none
      },
      align: if alignment in cel {align-horizontal-transforms.at(cel.alignment.horizontal) + align-vertical-transforms.at(cel.alignment.vertical)} else {left+bottom},
      {
      
      let rot = if alignment in cel and "rot" in  cel.alignment {cel.alignment.rot * -1deg} else {0deg}
      
      align(
        if alignment in cel {align-horizontal-transforms.at(cel.alignment.horizontal) + align-vertical-transforms.at(cel.alignment.vertical)} else {left+bottom},
        box(
          inset:0pt,
          //stroke:red+0.8pt,
          width: colwidths.at(cel.x - 1),
          height: rowHeights.at(cel.y - 1) - default-padding,
          clip: true,
          [#box(inset:0.2em,rotate(rot,cel.value, reflow: true))],
        )
        )
      },
    ))
  )
}


#let xlsxBytes = read("Book.xlsx", encoding: none);
#let myxlsx = parse-sheet(xlsxBytes, "Sheet1")

Here is the raw data collected from the Excel file:

#figure(
  caption: [Rendered Table],
  text(size: 6pt)[
    #rect(columns(3)[#align(left)[#myxlsx.json.cells]])],
)

The rendered Excel file looks like this:

#figure(
  caption: [Rendered Table],
  render-table(myxlsx, range: "A1:E20", use-boarders: true, use-merge: true),
)


#pagebreak()
#let themesFile = read("ThemeColors.xlsx", encoding: none);
#let colorSheet = parse-sheet(themesFile, "Sheet1")

#figure(
  caption: [Colors Table],
    render-table(colorSheet, range: none,scale:0.45),
)


#pagebreak()

Here we are only extracting the values using excel_to_json_values.

#let myxlsxv = parse-sheet-data(xlsxBytes, "Sheet1")

#table(columns: calc.max(..myxlsxv.map(x=>x.x)),
  ..myxlsxv.map(x=>x.value))
