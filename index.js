const xlsx = require('xlsx');
const path = require('path');
const firebase = require('firebase');
const createLineup = require('./lineup');
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

const uploadData = (year, gender) => {
  const db = firebase.database().ref(`lineupData/${gender}`);
  //read in the excel spreadsheet
  const file = xlsx.readFile(
    path.join(
      __dirname,
      `../../../BSD/Basketball/Tracker Data/${year}${
        gender === 'women' ? 'w' : ''
      }.xlsx`
    )
  );
  const results = {
    Total: {
      lineups: {},
      players: {},
      count: 0,
      yearlyTotal: createLineup(year),
    },
    Conference: {
      lineups: {},
      players: {},
      count: 0,
      yearlyTotal: createLineup(year),
    },
    ['Non-Conference']: {
      lineups: {},
      players: {},
      count: 0,
      yearlyTotal: createLineup(year),
    },
    Home: {lineups: {}, players: {}, count: 0, yearlyTotal: createLineup(year)},
    Away: {lineups: {}, players: {}, count: 0, yearlyTotal: createLineup(year)},
    Neutral: {
      lineups: {},
      players: {},
      count: 0,
      yearlyTotal: createLineup(year),
    },
    Q1: {lineups: {}, players: {}, count: 0, yearlyTotal: createLineup(year)},
    Q2: {lineups: {}, players: {}, count: 0, yearlyTotal: createLineup(year)},
    Q3: {lineups: {}, players: {}, count: 0, yearlyTotal: createLineup(year)},
    Q4: {lineups: {}, players: {}, count: 0, yearlyTotal: createLineup(year)},
  };
  const gamesList = [];
  file.SheetNames.forEach((game, index) => {
    //filter out the sheet metadata
    const {'!ref': ref, '!margins': margins, ...data} = file.Sheets[game];
    const gameLineups = {};
    const players = {};
    const keys = Object.keys(data);
    //grab and store the final 4 info cells
    const quad = data[keys.pop()].v;
    const accGame = data[keys.pop()].v;
    const gameLocation = game.includes('(N)')
      ? 'Neutral'
      : game.includes('@')
      ? 'Away'
      : 'Home';
    const oppScore = data[keys.pop()].v;
    const wakeScore = data[keys.pop()].v;

    //set the counts for the games
    results.Total.count++;
    results[accGame ? 'Conference' : 'Non-Conference'].count++;
    results[gameLocation].count++;
    results[`Q${quad}`].count++;

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
      //add the lineup to the array for the individual game
      gameLineups[lineup.players] = lineup;
      addPlayer(lineup.players, players, lineup);
      //add the lineup data to each grouping
      addLineup(results, lineup, accGame, gameLocation, quad);
      //update the yearly total
    }
    gamesList.push({
      game: game,
      order: index + 9,
      score: {
        opp: oppScore,
        wake: wakeScore,
      },
      lineups: gameLineups,
      players: players,
      gameCount: 0,
    });
  });
  //convert the data into 1 big array
  const dataArray = [];
  const yearlyTotal = [];
  Object.keys(results).forEach((category, i) => {
    dataArray.push({
      game: category,
      order: i,
      score: {
        opp: 0,
        wake: 0,
      },
      lineups: results[category].lineups,
      players: results[category].players,
      gameCount: results[category].count,
    });
    yearlyTotal.push({
      category: category,
      order: i,
      lineup: results[category].yearlyTotal,
    });
  });
  dataArray.push(...gamesList);

  //upload to the database
  return new Promise((res, rej) => {
    firebase
      .auth()
      .signInWithEmailAndPassword(process.env.EMAIL, process.env.PASSWORD)
      .then((userCredential) => {
        db.child(`${year}`)
          .set(dataArray)
          .then(() => {
            console.log(`Lineups uploaded for ${year} ${gender}`);
            db.child(`years/${year}`)
              .set(yearlyTotal)
              .then(() => {
                console.log(`Yearly totals uploaded for ${year} ${gender}`);
                res(`${year} ${gender} complete!`)
              });
          });
      });
  });
};
const addLineup = (res, lineup, conference, location, quad) => {
  const players = lineup.players;
  const conf = conference ? 'Conference' : 'Non-Conference';
  const q = `Q${quad}`;
  for (const key of ['Total', conf, location, q]) {
    if (res[key].lineups[players])
      combineLineups(res[key].lineups[players], lineup);
    else res[key].lineups[players] = {...lineup};
    combineLineups(res[key].yearlyTotal, {...lineup});
    addPlayer(players, res[key].players, {...lineup});
  }
};

const combineLineups = (parent, child) => {
  const keys = Object.keys(parent);
  keys.forEach((property) => {
    const childVal = child[property] ? child[property] : 0;
    if (typeof childVal === 'number') {
      parent[property] += childVal;
    }
  });
};

//add the lineup stats to all players in the lineup
const addPlayer = (players, results, lineup) => {
  const playerArray = players.split('\\');
  for (const player of playerArray) {
    if (!results[player]) results[player] = createLineup(`${player}`);
    combineLineups(results[player], lineup);
  }
};

const go = async () => {
  //Default year
  let year = '2025-26';
  if (process.argv[2] && process.argv[2] === 'all') {
    //upload all years
    const yearList = ['2020-21', '2021-22', '2022-23', '2023-24', '2024-25'];
    for (const season of yearList) {
      const status = await uploadData(season, 'men');
      console.log(status)
      const status2 = await uploadData(season, 'women');
      console.log(status2)
    }
    process.exit(0)
  } else {
    year =
      process.argv[2] && /[\d]{4}-[\d]{2}/.test(process.argv[2])
        ? process.argv[2]
        : year;
    const gender =
      process.argv[2] && process.argv[2].includes('w') ? 'women' : 'men';
    await uploadData(year.replace('w', ''), gender);
    process.exit();
  }

};
go()
