'use strict';
var start = new Date();
var XLSX = require('xlsx');
var moment = require('moment');
var fs = require('fs');
var connection;
var oracledb = require('oracledb');
const dbconfig = { user: "minimes_ff_wbr", password: "Baza0racl3appl1cs", connectString: "172.22.8.47/ORA" };


//NAZWY POLSKICH ZNAKOW PODAWAÆ JAKO \uKODZNAKU
var str = 'C:\\aaa\u0144\\tes t.xlsm';

console.clear();
async function ildPut() {
    var weeknumber;

    //TRYB_TEST
    var startDay = moment('2019-01-01');

    //TRYB NORMAL
    startDay = moment().subtract(1, 'day');
    var endDay = moment();


    //TESTOWANIE
   // console.log(endDay.diff(startDay, 'minutes'));
   // return;


    const connection = await oracledb.getConnection({
        user: "minimes_ff_wbr",
        password: "Baza0racl3appl1cs",
        connectString: "172.22.8.47/ORA"
    });

    for (var day = startDay; day < endDay; day.add(1, 'day')) {

        //day = moment("02-" + d + "-2019", "MM-DD-YYYY");
        weeknumber = day.isoWeek();

        str = '\\\\172.22.6.130\\pf$\\PW4 Wulkanizacja\\dzie\u0144\\ANALIZY\\\Analiza zdania stycze\u0144 2019\\Analiza zdania WEEK ' + (weeknumber < 10 ? '0' : '') + weeknumber + ' (' + day.format('YYYY.MM.DD') + ').xlsm';
        //console.log(str);
        try {
            if (fs.existsSync(str)) {
                //file exists
                var workbook = XLSX.readFile(str);
                if (workbook) {
                    var worksheet = workbook.Sheets["ILD_WBR"];
                    var ok = 0;
                    var okCure = 0;
                    var totalCure = 0;
                    var total = 0;
                    var i = 13;
                    while (worksheet['A' + i]) {

                        //KALKULACJA ILD CURE
                        if (worksheet['BD' + i]) {
                            if (worksheet['BS' + i]) {
                                if (worksheet['BS' + i].v == 'OK')
                                    okCure++;
                                totalCure++;
                            }

                        }

                        //KALKULACJI ILD STOCK
                        if (worksheet['Q' + i]) {
                            if (worksheet['Q' + i].v == 'OK')
                                ok++;
                            total++;
                        }
                        i++;

                    }

                }
                console.log('Analiza zdania WEEK ' + (weeknumber < 10 ? '0' : '') + weeknumber + ' (' + day.format('YYYY.MM.DD') + ').xlsm ILD=' + (total ? ((ok / total) * 100).toFixed(2) : 0) + ', ILD_CURE:' + (totalCure ? ((okCure / totalCure) * 100).toFixed(2) : 0));
                await connection.execute("update ff_data set ild=" + (total ? ((ok / total) * 100).toFixed(2) : 0) + ", ild_cure=" + (totalCure ? ((okCure / totalCure) * 100).toFixed(2) : 0)+ " where zmiana='TOTAL' and doba=to_date('" + day.format('YY-MM-DD') + "','yy-mm-dd')",
                    {}, //WIAZANIE ZMIENNYCH

                    {
                        //  resultSet: true maxRows: 1000000
                        // autoCommit: true
                    },
                    function (err, result) {
                        if (err) {
                            console.error(err);
                            return;
                        }
                        //console.log('WSTAWIONO');

                    });

            }
            else {
                console.log('BRAK PLIKU W LOKALIZACJI: ' + str);
            }
        }
        catch (err) {
            console.error(err);
        }


    }
    await connection.commit();
    await connection.close();
    var end = new Date() - start;
    console.info('Execution time: %dms', end);
}

ildPut();



