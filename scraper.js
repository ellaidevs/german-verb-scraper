const axios = require('axios')
const cheerio = require('cheerio')
const ExcelJS = require('exceljs')

const url = 'https://languageposters.com/pages/german-verbs'

let workbook = new ExcelJS.Workbook()
let worksheet = workbook.addWorksheet('Conjugations')

worksheet.columns = [
  { header: 'Conjugation', key: 'conjugation', width: 25 },
  { header: 'ich', key: 'ich', width: 25 },
  { header: 'du', key: 'du', width: 25 },
  { header: 'er-sie-es', key: 'er-sie-es', width: 25 },
  { header: 'wir', key: 'wir', width: 25 },
  { header: 'ihr', key: 'ihr', width: 25 },
  { header: 'sie-Sie', key: 'sie-Sie', width: 25 },
]

axios(url)
  .then((response) => {
    const html = response.data
    const $ = cheerio.load(html)
    const elements = $('div.rte--nomargin ul li')

    const promises = Array.from(elements).map((element, i) => {
      const $element = $(element)
      const aText = $element.find('a').text()
      const spanText = $element.find('span').text()
      const combinedText = aText + spanText
      const link = $element.find('a').attr('href')

      return axios(link)
        .then((response) => {
          const html = response.data
          const $ = cheerio.load(html)
          const rows = $('tbody tr')

          let conjugationObject = {
            conjugation: combinedText,
          }

          const pronouns = ['ich', 'du', 'er-sie-es', 'wir', 'ihr', 'sie-Sie']

          rows.each((index, row) => {
            if (index < 6) {
              const value = $(row).find('td:nth-child(2)').text()
              conjugationObject[pronouns[index]] = value
            }
          })

          return conjugationObject
        })
        .catch(console.error)
    })

    return Promise.all(promises)
  })
  .then((rows) => {
    rows.forEach((row) => {
      worksheet.addRow(row)
    })
    return workbook.xlsx.writeFile('Conjugations.xlsx')
  })
  .then(() => console.log('Excel file created successfully.'))
  .catch(console.error)
