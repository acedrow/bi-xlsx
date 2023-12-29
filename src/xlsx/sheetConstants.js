const TITLE_ROW_BACKGROUND_COLOR_ARGB = 'ffabcb95'

// the first row on the results sheet that we show data:
const RESULTS_TEAM_ROW_START = 7
// the first row on the pools sheet that we show data:
const POOLS_TEAM_ROW_START = 3
// the first row on the brackets sheet that we show data:
const BRACKETS_TEAM_ROW_START = 3

// how many 'fight rows' to pre-populate:
const NUMBER_FIGHTS_POOLS = 120

const TOURNAMENT_TYPE_TEAMS = 'teams'
const TOURNAMENT_TYPE_FIGHTERS = 'fighters'

const COL_SIZE = {
  xxs: 3.5,
  xs: 4.5,
  s: 6,
  m: 10,
  l: 15,
  xl: 21
}

const RESULTS_COLS = [
  { key: 'A', width: COL_SIZE.l },
  { key: 'B', width: COL_SIZE.xl },
  { key: 'C', width: COL_SIZE.s },
  { key: 'D', width: COL_SIZE.xxs },
  { key: 'E', width: COL_SIZE.xxs },
  { key: 'F', width: COL_SIZE.xxs },
  { key: 'G', width: COL_SIZE.xxs },
  { key: 'H', width: COL_SIZE.xxs },
  { key: 'I', width: COL_SIZE.xxs },
  { key: 'J', width: COL_SIZE.xxs },
  { key: 'K', width: COL_SIZE.xxs },
  { key: 'L', width: COL_SIZE.xxs },
  { key: 'M', width: COL_SIZE.xxs },
  { key: 'N', width: COL_SIZE.xxs },
  { key: 'O', width: COL_SIZE.xxs },
  { key: 'P', width: COL_SIZE.xxs },
  { key: 'Q', width: COL_SIZE.xxs },
  { key: 'R', width: COL_SIZE.xxs },
  { key: 'S', width: COL_SIZE.xxs },
  { key: 'T', width: COL_SIZE.s },
  { key: 'U', width: COL_SIZE.s },
  { key: 'V', width: COL_SIZE.s },
  { key: 'W', width: COL_SIZE.m },
  { key: 'X', width: COL_SIZE.m },
  { key: 'Y', width: COL_SIZE.m },
  { key: 'Z', width: COL_SIZE.l }
]

const POOLS_BRACKETS_COLS = [
  { key: 'A', width: COL_SIZE.s },
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
  { key: 'S', width: COL_SIZE.xs }
]

const POOL_BRACKET_HEADER_VALUES = {
  A: 'Fight',
  B: 'Team/Fighter ID',
  C: 'Team',
  D: 'Rounds Score',
  I: 'Fight',
  K: 'Rounds',
  O: 'Active/Grounded',
  R: 'Penalties'
}

const POOL_BRACKET_SUB_HEADER_VALUES = {
  D: 'R1',
  E: 'R2',
  F: 'R3',
  G: 'R4',
  H: 'R5',
  I: 'Fw',
  J: 'Fl',
  K: 'Rw',
  L: 'Rd',
  M: 'Rl',
  N: 'R',
  O: 'A',
  P: 'G',
  Q: 'A-G',
  R: 'YK',
  S: 'RK'
}

const getFont = (size, bold) => {
  return {
    name: 'Cambria',
    color: { argb: '00000000' },
    family: 2,
    size: size ?? 11,
    italic: false,
    bold: !!bold
  }
}

const TITLE_ROW_PROPS = {
  fill: {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: TITLE_ROW_BACKGROUND_COLOR_ARGB }
  },
  alignment: { horizontal: 'center' }
}

const EVENT_TIERS = [
  'Classic', 'Regional', 'Conference'
]

const FINALS_TYPE = [
  'Round Robin', 'Bracket'
]

const TEST_DATA = {
  eventName: 'test event',
  date: '1/2/3',
  location: 'St. Paul',
  tournamentType: 'teams',
  teams: [
    { name: 'wyverns', id: 'tha best' },
    { name: 'dfc', id: 'tha' },
    { name: 'knyaz', id: 'big boiz' }
  ],
  fighters: [
    { name: 'Derek', id: 'doc' },
    { name: 'Keegan', id: 'too big' },
    { name: 'Linden', id: 'lich boi' },
    { name: 'Vish', id: '420' },
    { name: 'Spence', id: 'coach' }
  ]
}
export {
  RESULTS_COLS,
  POOLS_BRACKETS_COLS,
  TITLE_ROW_PROPS,
  getFont,
  EVENT_TIERS,
  FINALS_TYPE,
  RESULTS_TEAM_ROW_START,
  POOLS_TEAM_ROW_START,
  BRACKETS_TEAM_ROW_START,
  NUMBER_FIGHTS_POOLS,
  TOURNAMENT_TYPE_TEAMS,
  TOURNAMENT_TYPE_FIGHTERS,
  TEST_DATA,
  POOL_BRACKET_SUB_HEADER_VALUES,
  POOL_BRACKET_HEADER_VALUES
}
