'use strict';
var start = new Date();
var XLSX = require('xlsx');
var moment = require('moment');
var fs = require('fs');
var connection;
var oracledb = require('oracledb');
const dbconfig = { user: "minimes_ff_wbr", password: "Baza0racl3appl1cs", connectString: "172.22.8.47/ORA" };


//NAZWY POLSKICH ZNAKOW PODAWAĆ JAKO \uKODZNAKU
var str = 'C:\\aaa\u0144\\tes t.xlsm';

console.clear();
    var weeknumber;
    //TRYB_TEST
    var startDay = moment('2019-03-01');
    var endDay = moment('2019-03-05');

    //TRYB NORMAL
    startDay = moment().subtract(2, 'day');
    endDay = moment();


    //TESTOWANIE
    // console.log(endDay.diff(startDay, 'minutes'));
    // return;


var shiftsArray = ['0321',
                '0321',
                '0321',
                '1032',
                '1032',
                '2103',
                '2103',
                '3210',
                '3210',
                '3210',
                '0321',
                '0321',
                '1032',
                '1032',
                '2103',
                '2103',
                '2103',
                '3210',
                '3210',
                '0321',
                '0321',
                '1032',
                '1032',
                '1032',
                '2103',
                '2103',
                '3210',
                '3210'
];

        var dept = 'BT4';

        fs.unlink('C:\\brama\\' + dept + '.txt', function (err) {
            if (err) throw err;
            console.log('File deleted!');
        });
        fs.appendFile('C:\\brama\\'+dept+'.txt', '', function (err) {
            if (err) throw err;
            console.log('Saved!');
        });


        str = 'C:\\brama\\'+dept+'.xlsx';
        //console.log(str);
        try {
            if (fs.existsSync(str)) {
                //file exists
                var workbook = XLSX.readFile(str);
                if (workbook) {
                    var worksheet = workbook.Sheets[dept];
                    var ok = 0;
                    var okCure = 0;
                    var totalCure = 0;
                    var total = 0;
                    var i = 2;
                    var j = 0;
                    //console.log(worksheet['A' + i].v.substr(14, 2));
                    while (worksheet['A' + i]) {
                        
                        //KALKULACJA ILD CURE
                        if (worksheet['C' + i] && worksheet['A' + i].v.substr(14, 2) == 'WY') {
                            var shiftDay = moment(worksheet['C' + i].v + ' ' + worksheet['D' + i].v);
                            var h = shiftDay.hour();
                            var shift;

                            if (h >= 6 && h < 14)
                                shift = 1;
                            else if (h >= 14 && h < 22)
                                shift = 2;
                            else
                                shift = 3;

                            var masterDay = moment('2016-07-02 00:00:00', 'YYYY-MM-DD HH:mm:ss');
                            shiftDay.subtract(6, 'hour');

                            var x = shiftDay.diff(masterDay, 'days')%28;
                            var brygada = worksheet['I' + i].v.substr(5, 1);

                            var shiftDec = shiftsArray[x].substr(brygada-1,1);

                            if ((h == 5 || h == 13 || h == 21) && shiftDec == shift) {
                                var tim;
                                if(h==5)
                                    tim = moment(worksheet['C' + i].v+' 06:00:00','YYYY-MM-DD HH:mm:ss');
                                if (h == 13)
                                    tim = moment(worksheet['C' + i].v + ' 14:00:00', 'YYYY-MM-DD HH:mm:ss');
                                if (h == 21)
                                    tim = moment(worksheet['C' + i].v + ' 22:00:00', 'YYYY-MM-DD HH:mm:ss');
                                
                                var out = moment(worksheet['C' + i].v + " " + worksheet['D' + i].v, 'YYYY-MM-DD HH:mm:ss');

                                var diff = tim.diff(out, 'seconds');
                               // console.log(shiftDay.format('YYYY-MM-DD HH:mm:ss') + ' ' + masterDay.format('YYYY-MM-DD HH:mm:ss') + ' Brygada:' +brygada+' zmiana:'+shiftDec);
                               // console.log(tim);

                                console.log(worksheet['A' + i].v+','+worksheet['G' + i].v + "," + worksheet['C' + i].v + "," + worksheet['D' + i].v + ","+shift+","+worksheet['J' + i].v + ","+diff+','+shiftDec);
                                fs.appendFile('C:\\brama\\' + dept + '2.txt', worksheet['A' + i].v + ',' + worksheet['G' + i].v + "," + worksheet['C' + i].v + "," + worksheet['D' + i].v + "," + shift + "," + worksheet['J' + i].v + "," + diff + ',' + shiftDec + "\n\r" , function (err) {
                                    if (err) throw err;
                                    //console.log('Updated!');
                                });
                                j++;
                            }

                        }

                        i++;

                    }
                }
            }
            else {
                console.log('BRAK PLIKU W LOKALIZACJI: ' + str);
            }
        }
        catch (err) {
            console.error(err);
        }

console.log(j);
    var end = new Date() - start;
    console.info('Execution time in node JS: %d second', end / 1000);






