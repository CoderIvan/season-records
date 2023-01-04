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
						}
					}
					return {
						winner: { team: teamB, score: scoreB },
						loser: { team: teamA, score: scoreA },
					}
				}),
		}))
}

function stats(dayRecords) {
	const doubles = {}
	const singles = {}

	dayRecords.forEach((dayReocrd) => {
		dayReocrd.records.forEach((record) => {
			const winTeam = record.winner.team
			const loseTeam = record.loser.team
			const netScore = record.winner.score - record.loser.score

			if (!doubles[winTeam]) {
				doubles[winTeam] = { win: 0, lose: 0, score: 0, netScore: 0, total: 0 }
			}
			doubles[winTeam].win += 1
			doubles[winTeam].score += record.winner.score
			doubles[winTeam].netScore += netScore
			doubles[winTeam].total += 1

			if (!doubles[loseTeam]) {
				doubles[loseTeam] = { win: 0, lose: 0, score: 0, netScore: 0, total: 0 }
			}
			doubles[loseTeam].lose += 1
			doubles[loseTeam].score += record.loser.score
			doubles[loseTeam].netScore -= netScore
			doubles[loseTeam].total += 1

			winTeam.split('').forEach((winner) => {
				if (!singles[winner]) {
					singles[winner] = { win: 0, lose: 0, score: 0, netScore: 0, total: 0 }
				}
				singles[winner].win += 1
				singles[winner].score += record.winner.score
				singles[winner].netScore += netScore
				singles[winner].total += 1
			})
			loseTeam.split('').forEach((loser) => {
				if (!singles[loser]) {
					singles[loser] = { win: 0, lose: 0, score: 0, netScore: 0, total: 0 }
				}
				singles[loser].lose += 1
				singles[loser].score += record.loser.score
				singles[loser].netScore -= netScore
				singles[loser].total += 1
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

	const statsObject = stats(dayRecords)
	console.log(statsObject)
}

parse(path.join(__dirname, './source/202210-202212/index.xlsx'))
