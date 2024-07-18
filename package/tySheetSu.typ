#let ss = plugin("tySheetSu.wasm")

#let parse-sheet(xlsxBytes, sheetname) = {
  let xlsx = (json: json.decode(ss.excel_to_json(xlsxBytes, bytes(sheetname))))
  return xlsx
}

#let parse-sheet-data(xlsxBytes, sheetname) = {
  let xlsx = (
    json: json.decode(ss.excel_to_json_values(xlsxBytes, bytes(sheetname))),
  )
  return xlsx
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
#let xlsx-draw-table(
  xlsx,
  range: none,
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

  )
  

  // Use excel scaling (this is hacky)
  let weirdExcelScaleX = 5 / 8.43 * 1em
  let weirdExcelScaleY = 5 / 38.4 * 1em
  let weridExcelRowHeightDefault = 14.4
  let weirdExcelColWidthDefault = 8.75

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
    .map(x => if (x.width == 0 and x.hidden == false) {
    weirdExcelColWidthDefault * weirdExcelScaleX
  } else if x.hidden==true { 0pt}
  else {
    x.width * weirdExcelScaleX
  })

  //Set row heights based on excel input using odd number for default
  let rowHeights = range(1,rows + 1)
    .map(y => xlsx.json.rows.filter(row => row.y == y).at(
        0,
        default: (height: weridExcelRowHeightDefault),
      ))
    .map(x => if (x.height == 0 and x.hidden == false) {
    weridExcelRowHeightDefault * weirdExcelScaleY
  } else if x.hidden==true { 0pt}
  else {
    x.height * weirdExcelScaleY
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
        color.rgb("#" + cel.fill_color.slice(-6))
      } else {
        none
      },
      {

      let text-align = if alignment in cel {align-horizontal-transforms.at(cel.alignment.horizontal)} else {left}
      box(
        inset:0pt,
        //stroke:red+0.8pt,
        width: colwidths.at(cel.x - 1),
        height: rowHeights.at(cel.y - 1) - default-padding,
        clip: true,
        align(text-align)[#box(inset:0.2em,cel.value)],
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
  xlsx-draw-table(myxlsx, range: "A1:E20", use-boarders: true),
)
 
#pagebreak()

Here we are only extracting the values using excel_to_json_values.

#let myxlsxv = parse-sheet-data(xlsxBytes, "Sheet1")

#myxlsxv

#let sheet-data-as-array(sheet-data) = {
  let a = sheet-data.json.cells
  a.sorted(key: q => q.x).sorted(key: q => q.y).map(x => x.value)
}

#sheet-data-as-array(myxlsxv)

#table(columns: 6,
  ..sheet-data-as-array(myxlsxv))