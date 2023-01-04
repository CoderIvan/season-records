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
					let [teamA, , , teamB] = cell
					const [, scoreA, scoreB] = cell
					teamA = trim(teamA)
					teamB = trim(teamB)
					if (scoreA > scoreB) {
						return {
							winner: { team: teamA, score: scoreA },
							loser: { team: teamB, score: scoreB },
							isDeuce: scoreA > 21,
						}
					}
					return {
						winner: { team: teamB, score: scoreB },
						loser: { team: teamA, score: scoreA },
						isDeuce: scoreB > 21,
					}
				}),
		}))
}

function newRecord(object, key, isWin, isDeuce, score, netScore) {
	if (!object[key]) {
		object[key] = { win: 0, lose: 0, score: 0, netScore: 0, total: 0, deuce: 0 }
	}
	if (isWin) {
		object[key].win += 1
	} else {
		object[key].lose += 1
	}
	object[key].score += score
	if (isWin) {
		object[key].netScore += netScore
	} else {
		object[key].netScore -= netScore
	}
	if (isDeuce) {
		object[key].deuce += 1
	}
	object[key].total += 1
}

function stats(dayRecords) {
	const doubles = {}
	const singles = {}

	dayRecords.forEach((dayReocrd) => {
		dayReocrd.records.forEach(({ winner, loser, isDeuce }) => {
			const netScore = winner.score - loser.score

			;[winner, loser].forEach(({ team, score }, index) => {
				const isWin = index === 0
				newRecord(doubles, team, isWin, isDeuce, score, netScore)
				team.split('').forEach((parter) => {
					newRecord(singles, parter, isWin, isDeuce, score, netScore)
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
			console.log('%j', record)
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
				console.log('%s %s', key, statsObject[keys][key])
			})
		console.log('-------------------------------------')
	})
	console.log('-------------------------------------')
}

parse(path.join(__dirname, './source/202210-202212/index.xlsx'))
