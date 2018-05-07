var xlsx = require('node-xlsx');
var fs = require("fs");
var path = require("path")
processExcel()
function processExcel() {
    var obj = xlsx.parse(__dirname + '/test.xls');
    var sheets = obj
    var sheet = sheets[0]
    let sheetName = sheet.name
    let sheetData = sheet.data
    // console.log(sheetName)
    let lastDate
    let isWuye
    var rawData = []
    var rawPeer = []
    for (let dataLine of sheetData) {

        if ((dataLine instanceof Array)) {
            // console.log(dataLine)
            let isDate = false
            let isTitle = false
            let firstData = dataLine[0]

            if (firstData != undefined) {
                let index = firstData.indexOf("2018-")
                if (index >= 0) {
                    lastDate = firstData.substring(5, 10)
                    isDate = true
                }
                if (firstData == "客户名称") {
                    isTitle = true
                }
                if (firstData == '合顺物业') {
                    isWuye = true
                } else if (firstData == '合顺物业/酒水') {
                    isWuye = false
                }
            }

            if (isDate || isTitle) {
                // console.log(firstData)
            } else {
                //真是数据部分
                dataLine['date'] = lastDate
                dataLine['month'] = Number(lastDate.substring(0, 2))
                dataLine['day'] = Number(lastDate.substring(3, 5))
                dataLine.shift()
                if (isWuye) {
                    rawData.push(dataLine)
                } else {
                    rawPeer.push(dataLine)
                }
            }
        }
    }

    let dataWuyue = processSheet(rawData)
    let dataJiushui = processSheet(rawPeer)
    let data = [{
        name: '物业',
        data: dataWuyue
    }, {
        name: '酒水',
        data: dataJiushui
    }]
    writeExcel(data)
}


function processSheet(rawData) {
    var retData = new Array()
    for (let data of rawData) {
        let date = data['date']
        let day = data['day']
        let month = data['month']
        let name = data[0]
        let danwei = data[1]
        let count = data[2]
        let price = data[3]
        let money = data[4]

        let item = {}
        item.day = day
        item.month = month
        item.name = name
        item.danwei = danwei
        item.count = count
        item.price = price
        item.money = money
        item.date = date

        let exist = retData[name]
        if (exist == undefined) {
            retData[name] = [item]
        } else {
            retData[name].push(item)
        }
    }


    let days = []

    for (let i = 1; i < 31; i++) {
        let find = false
        for (let key in retData) {
            let tempData = retData[key]
            for (let ttmp of tempData) {
                let day = ttmp.day
                if (i == day) {
                    find = true
                    break
                }
            }

        }
        if (find) {
            days.push(i)
        }
    }


    let rowList = []
    let row1 = ['']
    let row2 = ['商品列表']
    rowList.push(row1)
    rowList.push(row2)
    for (let day of days) {
        row1.push(day + '日')
        row1.push('')
        row1.push('')
        row2.push('单价')
        row2.push('数量')
        row2.push('总额')
    }

    //名字进行排序
    let keys = new Array()
    for (let key in retData) {
        keys.push(key)
    }
    keys.sort()

    for (let key of keys) {
        let valueArray = retData[key]
        let name = key
        if(name == '五花肉（整片）'){
            let a = 10
        }
        let temprowlist = []
        let dayIndex = -1
        for (let tempday of days) {
            dayIndex++
            let index = 0 
            for (let value of valueArray) {
                let day = value.day
                if (day == tempday) {
                    let tempRow = null
                    if(temprowlist.length > index){
                        tempRow = temprowlist[index]
                    }else{
                        tempRow = []
                        temprowlist.push(tempRow)
                        tempRow.push(name)
                    }

                    while(tempRow.length-1 < dayIndex * 3){
                        tempRow.push('')
                        tempRow.push('')
                        tempRow.push('')
                    }


                    index++

                    tempRow.push(value.count)
                    tempRow.push(value.price)
                    tempRow.push(value.money)
                }
            }
        }
        for(let tempRow of temprowlist){
            rowList.push(tempRow)
        }
    }
    return rowList
}

function writeExcel(data) {
    var buffer = xlsx.build(data);
    fs.writeFile('./resut.xlsx', buffer, function (err) {
        if (err) throw err;
        console.log('has finished');
    });
}
