const xlsx = require('xlsx');
const path = require('path');
const firebase = require('firebase');
require('dotenv').config();

//SET UP FIREBASE
firebase.initializeApp({
  apiKey: process.env.API_KEY,
  authDomain: process.env.AUTH_DOMAIN,
  databaseURL: process.env.DATABASE_URL,
  projectId: process.env.PROJECT_ID,
  storageBucket: process.env.STORAGE_BUCKET,
  messagingSenderId: process.env.MESSAGING_SENDER_ID,
  appId: process.env.APP_ID,
});

//Default year
let year = '2024-25';
const type = process.argv[2] && process.argv[2] === 'w' ? 'women' : 'men';
year =
  process.argv[3] && /[\d]{4}-[\d]{2}/.test(process.argv[3])
    ? process.argv[3]
    : year;
const w = type === 'women' ? 'w' : '';
const db = firebase.database().ref(`lineupData/${type}`);

const parseData = () => {
  //determine the year and which team we are we dealing with

  //read in the excel spreadsheet
  const file = xlsx.readFile(
    path.join(
      __dirname,
      `../../../BSD/Basketball/Tracker Data/${year}${w}.xlsx`
    )
  );
  const results = {
    total: new Map(),
    conference: new Map(),
    nonConference: new Map(),
    home: new Map(),
    away: new Map(),
    quad1: new Map(),
    quad2: new Map(),
    quad3: new Map(),
    quad4: new Map(),
    games: {},
  };
  file.SheetNames.forEach((game, i) => {
    //filter out the sheet metadata
    const {'!ref': ref, '!margins': margins, ...data} = file.Sheets[game];
    const lineups = {};
    const keys = Object.keys(data);
    //grab and store the final 4 info cells
    const quad = data[keys.pop()].v;
    const accGame = data[keys.pop()].v;
    const homeGame = game.includes('@') || game.includes('at') ? false : true;
    const oppScore = data[keys.pop()].v;
    const wakeScore = data[keys.pop()].v;
    //all that is left in array is keys of cells containing lineup data
    for (let i = 26; i < keys.length; i += 26) {
      const lineup = {
        players: data[keys[i]].v,
        time: data[keys[i + 1]].v,
        pointsFor: data[keys[i + 2]].v,
        pointsAgainst: data[keys[i + 3]].v,
        dRebFor: data[keys[i + 4]].v,
        dRebAgainst: data[keys[i + 5]].v,
        oRebFor: data[keys[i + 6]].v,
        oRebAgainst: data[keys[i + 7]].v,
        ftaFor: data[keys[i + 8]].v,
        ftaAgainst: data[keys[i + 9]].v,
        madeTwosFor: data[keys[i + 10]].v,
        attemptedTwosFor: data[keys[i + 11]].v,
        madeTwosAgainst: data[keys[i + 12]].v,
        attemptedTwosAgainst: data[keys[i + 13]].v,
        madeThreesFor: data[keys[i + 14]].v,
        attemptedThreesFor: data[keys[i + 15]].v,
        madeThreesAgainst: data[keys[i + 16]].v,
        attemptedThreesAgainst: data[keys[i + 17]].v,
        assistsFor: data[keys[i + 18]].v,
        assistsAgainst: data[keys[i + 19]].v,
        turnoversFor: data[keys[i + 20]].v,
        turnoversAgainst: data[keys[i + 21]].v,
        paintFor: data[keys[i + 22]].v,
        paintAgainst: data[keys[i + 23]].v,
        secondFor: data[keys[i + 24]].v,
        secondAgainst: data[keys[i + 25]].v,
      };
      lineups[lineup.players] = lineup;
      addLineup(results, lineup, accGame, homeGame, quad);
    }
    results.games[game] = {
      accGame: accGame,
      order: i,
      score: {
        opp: oppScore,
        wake: wakeScore,
      },
      quad: quad,
      lineups: lineups,
    };
  });

  return results;
};
const addLineup = (res, lineup, conference, home, quad) => {
  const players = lineup.players;

  //all lineups should be added to the total
  if (res.total.has(players)) combineLineups(res.total.get(players), lineup);
  else res.total.set(players, lineup);

  //add lineups to conference/noncon
  const conf = conference ? 'conference' : 'nonConference';
  if (res[conf].has(players)) combineLineups(res[conf].get(players), lineup);
  else res[conf].set(players, lineup);

  //add home or away
  const location = home ? 'home' : 'away';
  if (res[location].has(players))
    combineLineups(res[location].get(players), lineup);
  else res[location].set(players, lineup);

  //add quad
  const q = `quad${quad}`;
  if (res[q].has(players)) combineLineups(res[q].get(players), lineup);
  else res[q].set(players, lineup);
};

const combineLineups = (parent, child) => {
  const keys = Object.keys(parent);
  keys.forEach((property) => {
    const prop = child[property];
    if (typeof prop === 'number') {
      parent[property] += prop;
    }
  });
};

const data = parseData();
db.child(`test/${year}`)
  .set(data)
  .then(() => {
    console.log('Lineups Uploaded');
    process.exit(0);
  });
