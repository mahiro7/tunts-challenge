const { GoogleSpreadsheet } = require('google-spreadsheet')
const credentials = require('./credentials.json')

const docID = '1Ve0F3ROXiDMmpVDTCsky3X6xAMQCPyrCyr9ge04x6B8'
const doc = new GoogleSpreadsheet(docID)

const aveg = (p1, p2, p3) => (p1+p2+p3)/3/10 // average account function

const absent = (f, m) => f > 0.25*m ? true : false // function to verify absence

const situation = m => { // function to define student situation
    const s0 = 'Reprovado por Nota'
    const s1 = 'Exame Final'
    const s2 = 'Aprovado'

    if (m < 5){
        return s0
    }else if (m >= 5 && m < 7){
        return s1
    }else{
        return s2
    }
}

const aprov = m => 10-m

async function accessSpreadsheet() {
    await doc.useServiceAccountAuth({
        client_email: credentials.client_email,
        private_key: credentials.private_key.replace(/\\n/g, '\n')
    })

    await doc.loadInfo(); // loads document properties and worksheets
    const sheet = doc.sheetsByIndex[0]

    console.log(sheet.title)

    await sheet.loadCells('A2:H1000') // loads cells
    
    // update 'Situação' column
    var sita1 = []
    var sit = []
    var count = 0
    const x = ((sheet.cellStats.nonEmpty)-2)/6+2 // 
    for (var i = 3; i < x; i++){
        sita1[count] = sheet.getCell(i, 6).a1Address // get a1Address of 'Situação' col

        var p1 = parseInt(sheet.getCell(i, 3).formattedValue) // get data to average calc
        var p2 = parseInt(sheet.getCell(i, 4).formattedValue)
        var p3 = parseInt(sheet.getCell(i, 5).formattedValue)

        var faltas = parseInt(sheet.getCell(i, 2).formattedValue)
        const aulas = parseInt(sheet.getCell(1, 0).formattedValue.replace( /^\D+/g, ''))

        sit[count] = aveg(p1, p2, p3)

        if (absent(faltas, aulas) == true){
            sheet.getCellByA1(sita1[count]).value = 'Reprovado por Falta'
        } else {
        sheet.getCellByA1(sita1[count]).value = situation(sit[count])
        }

        console.log('Aluno: ',count,', Faltas: ',faltas,', P1: ',p1,', P2: ',p2,', P3: ',p3,', Média: ',sit[count])

        count++
    }
    await sheet.saveUpdatedCells()
    console.log('#######################')
    //////////
    //update 'Nota para Aprovação' column
    count = 0
    nota = []
    for(var i = 3; i < x; i++){
        nota[count] = sheet.getCell(i, 7).a1Address // get a1Address of 'Nota' column

        if (sheet.getCellByA1(sita1[count]).value != 'Exame Final'){
            sheet.getCellByA1(nota[count]).value = '0'
        }else{
            sheet.getCellByA1(nota[count]).value = Math.ceil(aprov(sit[count])*10)
        }
        
        count++
    }
    await sheet.saveUpdatedCells()
}
accessSpreadsheet()