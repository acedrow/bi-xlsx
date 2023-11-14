import { RESULTS_COLS, RESULTS_TEAM_ROW_START, TITLE_ROW_PROPS, getFont } from './sheetConstants.js'
// sheet 1 Results
const generateTitleRows = (worksheet) => {
// ROW 1
  worksheet.addRow({ A: 'BUHURT INTERNATIONAL' })

  worksheet.getCell('A1').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A1').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A1').font = getFont(16, true)
  // ROW 2
  worksheet.addRow({ A: 'Event Scoring Sheet' })

  worksheet.getCell('A2').fill = TITLE_ROW_PROPS.fill
  worksheet.getCell('A2').alignment = TITLE_ROW_PROPS.alignment
  worksheet.getCell('A2').font = getFont(12, true)

  worksheet.mergeCells('A1:Z1')
  worksheet.mergeCells('A2:Z2')
}

const generateHeaderRows = (worksheet, eventName, date, location) => {
  const row3 = worksheet.addRow({
    A: 'Event Name:',
    B: eventName,
    T: 'Event Date:',
    W: date,
    Y: 'Event Tier:',
    Z: 'Classic'
  })

  row3.font = getFont(11, true)
  worksheet.getCell('Z3').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"Classic,Regional,Conference"']
  }
  worksheet.mergeCells('B3:S3')
  worksheet.mergeCells('T3:V3')
  worksheet.mergeCells('W3:X3')

  const row4 = worksheet.addRow({
    A: 'Event Location:',
    B: location,
    T: 'Finals Type:',
    W: 'Round Robin',
    Y: 'Tier Mult.',
    Z: { formula: '=IF(Z3="Regional",1.5,IF(Z3="Conference",2,1))' }
  })

  row4.font = getFont(11, true)
  worksheet.getCell('W4').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"Round Robin,Bracket"']
  }
  worksheet.mergeCells('B4:S4')
  worksheet.mergeCells('T4:V4')
  worksheet.mergeCells('W4:X4')
}

const generateTeamHeaders = (worksheet) => {
  const headerFont = getFont(11, true)
  const wrapAlignment = { vertical: 'middle', horizontal: 'center', wrapText: false }
  // Add headers for row 5 with common font and alignment
  const row5 = worksheet.addRow({
    A: 'ID',
    B: 'Team',
    C: 'T',
    D: 'Fights',
    I: 'Rounds',
    N: 'Score',
    T: 'Cards',
    W: 'Points',
    X: 'Placement',
    Y: 'Rank adj. Points',
    Z: 'Final Awarded Points'
  })

  row5.font = headerFont
  row5.alignment = wrapAlignment

  // Add headers for row 6 and styles
  const row6 = worksheet.addRow({
    // Fights
    D: 'Won',
    E: 'Loss',
    F: 'Total',
    G: 'Win Ratio',
    // H: '', // will not be used but keep it ?
    // Rounds
    I: 'Win',
    J: 'Draw',
    K: 'Loss',
    L: 'Total',
    M: 'win ratio',
    // Active/Grounded
    N: 'Active',
    O: 'per round',
    P: 'Grounded',
    Q: 'per round',
    R: 'A/G difference', // active versus ground difference
    S: 'A/G Ratio', // active ground ratio read "kill/death Ratio"
    // penaties
    T: 'Yellow',
    U: 'Red',
    V: 'Total'
  })

  row6.font = headerFont
  row6.alignment = wrapAlignment

  worksheet.mergeCells('A5:A6')
  worksheet.mergeCells('B5:B6')
  worksheet.mergeCells('C5:C6')
  worksheet.mergeCells('D5:H5')
  worksheet.mergeCells('I5:M5')
  worksheet.mergeCells('N5:S5')
  worksheet.mergeCells('W5:W6')
  worksheet.mergeCells('X5:X6')
  worksheet.mergeCells('Y5:Y6')
  worksheet.mergeCells('Z5:Z6')
  worksheet.mergeCells('T5:V5')

  worksheet.getColumn('H').hidden = true
}

// generates the team names + id
const generateTeamDataRows = (worksheet, teams) => {
  teams.forEach((team, index) => {
    const rowIndex = index + RESULTS_TEAM_ROW_START
    const addedRow = worksheet.addRow({
      A: team.id,
      B: team.name,
      C: 1,
      D: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$I:$I)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$I:$I)` },
      E: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$J:$J)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$J:$J)` },
      F: { formula: `=SUM(D${rowIndex}:E${rowIndex})` },
      G: { formula: `=iferror(D${rowIndex}/F${rowIndex})` },
      H: { formula: `=F${rowIndex}/C${rowIndex}` },
      I: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$K:$K)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$K:$K)` },
      J: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$L:$L)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$L:$L)` },
      K: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$M:$M)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$M:$M)` },
      L: { formula: `=SUM(I${rowIndex}:K${rowIndex})` },
      M: { formula: `=iferror(I${rowIndex}/L${rowIndex})` },
      N: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$O:$O)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$O:$O)` },
      O: { formula: `=iferror(N${rowIndex}/L${rowIndex})` },
      P: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$P:$P)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$P:$P)` },
      Q: { formula: `=iferror(P${rowIndex}/L${rowIndex})` },
      R: { formula: `=N${rowIndex}-P${rowIndex}` },
      S: { formula: `=iferror(N${rowIndex}/(L${rowIndex}*5))` },
      T: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$R:$R)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$R:$R)` },
      U: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$S:$S)+SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$S:$S)` },
      V: { formula: `=T${rowIndex}+(2*U${rowIndex})` },
      W: { formula: `=SUMIF(pools!$B:$B,$A${rowIndex},pools!$I:$I)+(SUMIF(brackets!$B:$B,$A${rowIndex},brackets!$I:$I)*2)` },
      X: { formula: `=IF($W$4="Round Robin",RANK.EQ(D${rowIndex},$D$7:$D$100)+COUNTIFS($D$7:$D$100,D${rowIndex},$M$7:$M$100,">"&M${rowIndex})+COUNTIFS($D$7:$D$100,D${rowIndex},$M$7:$M$100,M${rowIndex},$S$7:$S$100,">"&S${rowIndex})+COUNTIFS($D$7:$D$100,D${rowIndex},$M$7:$M$100,M${rowIndex},$S$7:$S$100,S${rowIndex},$V$7:$V$100,"<"&V${rowIndex}),IFERROR(MATCH(A${rowIndex},brackets!$X$2:$X$5,0),"NA"))` },
      Y: { formula: `=IF(X${rowIndex}=3;W${rowIndex}+2;IF(X${rowIndex}=2;W${rowIndex}+4;IF(X${rowIndex}=1;W${rowIndex}+6;W${rowIndex})))` },
      Z: { formula: `=Y${rowIndex}*$Z$4` }
    })
    addedRow.font = getFont(11, false)
  }
  )
}

const generateResultsSheet = (workbook, { eventName, date, location, teams }) => {
  const resultsSheet = workbook.addWorksheet('results')
  resultsSheet.columns = RESULTS_COLS
  generateTitleRows(resultsSheet)
  generateHeaderRows(resultsSheet, eventName, date, location, teams)
  generateTeamHeaders(resultsSheet)
  generateTeamDataRows(resultsSheet, teams)
}

export default generateResultsSheet
