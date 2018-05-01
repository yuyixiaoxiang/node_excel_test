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


    let value = []
    let item = ['']
    let item1 = ['商品列表']
    value.push(item)
    value.push(item1)
    for (let day of days) {
        item.push(day + '日')
        item.push('')
        item.push('')
        item1.push('单价')
        item1.push('数量')
        item1.push('总额')
    }

    let keys = new Array()
    for (let key in retData) {
        keys.push(key)
    }
    keys.sort()

    for (let key of keys) {
        let valueArray = retData[key]
        let name = key
        let item = []
        item.push(name)
        for (let tempday of days) {
            let count
            let price
            let money
            for (let value of valueArray) {
                let day = value.day
                if (day == tempday) {
                    count = value.count
                    price = value.price
                    money = value.money
                    break
                }
            }
            item.push(count)
            item.push(price)
            item.push(money)
        }
        value.push(item)
    }
    return value
}

function writeExcel(data) {
    var buffer = xlsx.build(data);
    fs.writeFile('./resut.xlsx', buffer, function (err) {
        if (err) throw err;
        console.log('has finished');
    });
}
