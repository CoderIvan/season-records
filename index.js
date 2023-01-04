const Excel = require('exceljs')
const path = require('path')

function trim(value) {
	return value
		.replace(' ', '')
		.trim()
		.split('')
		.sort()
		.join('')
}

async function xlsxToArray(filename) {
	const workbook = await new Excel.Workbook().xlsx.readFile(filename)

	return workbook.worksheets
		.filter(({ name }) => /^\d{8}/.test(name))
		.map((worksheet) => ({
			name: worksheet.name.slice(0, 8),
			records: worksheet.getSheetValues().slice(2, 12)
				.map((row) => {
					const cell = row.slice(2, 6)
					const [teamA, scoreA, scoreB, teamB] = cell
					if (scoreA > scoreB) {
						return {
							winner: { team: trim(teamA), score: scoreA },
							loser: { team: trim(teamB), score: scoreB },
						}
					}
					return {
						winner: { team: trim(teamB), score: scoreB },
						loser: { team: trim(teamA), score: scoreA },
					}
				}),
		}))
}

function newRecord(object, key, isWin, winnerScore, loserScore) {
	if (!object[key]) {
		object[key] = { win: 0, lose: 0, score: 0, netScore: 0, total: 0, deuce: 0 }
	}
	if (isWin) {
		object[key].win += 1
	} else {
		object[key].lose += 1
	}
	object[key].score += isWin ? winnerScore : loserScore
	const netScore = winnerScore - loserScore
	if (isWin) {
		object[key].netScore += netScore
	} else {
		object[key].netScore -= netScore
	}
	if (winnerScore > 21) {
		object[key].deuce += 1
	}
	object[key].total += 1
}

function stats(dayRecords) {
	const doubles = {}
	const singles = {}

	dayRecords.forEach((dayReocrd) => {
		dayReocrd.records.forEach(({ winner, loser }) => {
			[winner, loser].forEach(({ team }, index) => {
				const isWin = index === 0
				newRecord(doubles, team, isWin, winner.score, loser.score)
				team.split('').forEach((parter) => {
					newRecord(singles, parter, isWin, winner.score, loser.score)
				})
			})
		})
	})

	return {
		doubles,
		singles,
	}
}

async function parse(filename) {
	const dayRecords = await xlsxToArray(filename)
	dayRecords.forEach((dayRecord) => {
		console.log(`-------------------${dayRecord.name} Start-------------------`)
		dayRecord.records.forEach((record) => {
			console.log(record)
		})
		console.log(`-------------------${dayRecord.name} End-------------------`)
	})

	console.log('-------------------------------------')
	const statsObject = stats(dayRecords)
	;['doubles', 'singles'].forEach((keys) => {
		Object
			.keys(statsObject[keys])
			.sort((keyA, keyB) => {
				if (statsObject[keys][keyB].win === statsObject[keys][keyA].win) {
					return statsObject[keys][keyB].netScore - statsObject[keys][keyA].netScore
				}
				return statsObject[keys][keyB].win - statsObject[keys][keyA].win
			}).forEach((key) => {
				const r = statsObject[keys][key]
				console.log(key, r, `${Math.floor((r.win / r.total) * 10000) / 100}%`)
			})
		console.log('-------------------------------------')
	})
	console.log('-------------------------------------')
}

parse(path.join(__dirname, './source/202210-202212/index.xlsx'))
