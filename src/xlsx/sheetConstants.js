const TITLE_ROW_BACKGROUND_COLOR_ARGB = 'ffabcb95'

const COL_SIZE = {
  xs: 3.5,
  s: 6,
  m: 10,
  l: 13,
  xl: 21
}

const RESULTS_COLS = [
  { key: 'A', width: COL_SIZE.l },
  { key: 'B', width: COL_SIZE.xl },
  { key: 'C', width: COL_SIZE.s },
  { key: 'D', width: COL_SIZE.xs },
  { key: 'E', width: COL_SIZE.xs },
  { key: 'F', width: COL_SIZE.xs },
  { key: 'G', width: COL_SIZE.xs },
  { key: 'H', width: COL_SIZE.xs },
  { key: 'I', width: COL_SIZE.xs },
  { key: 'J', width: COL_SIZE.xs },
  { key: 'K', width: COL_SIZE.xs },
  { key: 'L', width: COL_SIZE.xs },
  { key: 'M', width: COL_SIZE.xs },
  { key: 'N', width: COL_SIZE.xs },
  { key: 'O', width: COL_SIZE.xs },
  { key: 'P', width: COL_SIZE.xs },
  { key: 'Q', width: COL_SIZE.xs },
  { key: 'R', width: COL_SIZE.xs },
  { key: 'S', width: COL_SIZE.xs },
  { key: 'T', width: COL_SIZE.s },
  { key: 'U', width: COL_SIZE.s },
  { key: 'V', width: COL_SIZE.s },
  { key: 'W', width: COL_SIZE.m },
  { key: 'X', width: COL_SIZE.m },
  { key: 'Y', width: COL_SIZE.m },
  { key: 'Z', width: COL_SIZE.l }
]

const TITLE_ROW_PROPS = {
  font: {
    name: 'Cambria',
    color: { argb: '00000000' },
    family: 2,
    size: 16,
    italic: false,
    bold: true
  },
  fill: {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: TITLE_ROW_BACKGROUND_COLOR_ARGB }
  },
  alignment: {
    horizontal: 'center'
  }
}

export { RESULTS_COLS, TITLE_ROW_PROPS }
