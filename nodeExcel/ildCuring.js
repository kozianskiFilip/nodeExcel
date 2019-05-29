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
    var startDay = moment('2019-05-01');
    var endDay = moment('2019-05-28');

    //TRYB NORMAL
    //startDay = moment().subtract(2, 'day');
    //endDay = moment();


    //TESTOWANIE
    // console.log(endDay.diff(startDay, 'minutes'));
    // return;




    for (var day = startDay; day <= endDay; day.add(1, 'day')) {

        //day = moment("02-" + d + "-2019", "MM-DD-YYYY");
        weeknumber = day.isoWeek();

        str = '\\\\172.22.6.130\\pf$\\PW4 Wulkanizacja\\dzie\u0144\\ARCHIWUM - DOBA\\2019\\doba ' + day.format('YYYY.MM.DD') + '.xlsm';
        //console.log(str);
        try {
            if (fs.existsSync(str)) {
                //file exists
                var workbook = XLSX.readFile(str);
                if (workbook) {
                    var worksheet = workbook.Sheets["ASORTYMENTY_DOBA"];
                    var ok = 0;
                    var okCure = 0;
                    var totalCure = 0;
                    var total = 0;
                    var i = 4;

                    while (worksheet['C' + i]) {

                        if (worksheet['S' + i]) {
                            console.log(weeknumber + ' '+ day.format('YYYY.MM.DD')+ ' ' + worksheet['C' + i].v + ' ' + (worksheet['S' + i].v));
                        }
                        i++;
                    }
                }
               // console.log('Analiza z');
            }
            else {
                console.log('BRAK PLIKU W LOKALIZACJI: ' + str);
            }
        }
        catch (err) {
            console.error(err);
        }
    }

    var end = new Date() - start;
    console.info('Execution time in node JS: %d second', end / 1000);




