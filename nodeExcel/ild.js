'use strict';
var start = new Date();
var XLSX = require('xlsx');
var moment = require('moment');
var fs = require('fs');
var connection;
var oracledb = require('oracledb');
const dbconfig = { user: "minimes_ff_wbr", password: "Baza0racl3appl1cs", connectString: "172.22.8.47/ORA" };


//NAZWY POLSKICH ZNAKOW PODAWAÆ JAKO \uKODZNAKU
var str;

console.clear();

async function ildPut() {
    var weeknumber;

    //TRYB_TEST
    var startDay = moment('2019-05-26');
    var endDay = moment('2019-05-26');
    
    //TRYB NORMAL
    startDay = moment().subtract(1, 'day');
    endDay = moment();
    //endDay = moment().subtract(1, 'day');


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
        try {
            var items = fs.readdirSync('\\\\172.22.6.130\\pf$\\PW4 Wulkanizacja\\dzie\u0144\\ANALIZY\\');
            // console.log(items);
            for (var i = 0; i < items.length; i++) {
                str = 'sd';
                //console.log('(' + day.format('YYYY.MM.DD') + ').xlsm');
                if (items[i].indexOf('(' + day.format('YYYY.MM.DD') + ').xlsm') > 0) {
                    str = '\\\\172.22.6.130\\pf$\\PW4 Wulkanizacja\\dzie\u0144\\ANALIZY\\' + items[i];
                    console.log(str);
                    break;
                }
            } 
        } catch (err) {
            console.log(err);
        }



       // str = '\\\\172.22.6.130\\pf$\\PW4 Wulkanizacja\\dzie\u0144\\ANALIZY\\Analiza zdania WEEK ' + (weeknumber < 10 ? '0' : '') + (weeknumber) + ' (' + day.format('YYYY.MM.DD') + ').xlsm';
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

                    var insertIldString = 'INSERT ALL ';

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
                        insertIldString += "into ff_ild_plan(day, description,dpics,ctcode,sapcode,plan_week,plan_wtd,cure_wtd,stock_wtd,plan_today,wd11) values (trunc(sysdate,'ddd'),'" + worksheet['A' + i].v + "','" + worksheet['B' + i].v + "','" + worksheet['C' + i].v + "'," + worksheet['F' + i].v + ", " + worksheet['I' + i].v + "," + (worksheet['J' + i] ? worksheet['J' + i].v : 0) + "," + (worksheet['K' + i] ? worksheet['K' + i].v : 0) + "," + (worksheet['L' + i] ? worksheet['L' + i].v : 0) + "," + (worksheet['P' + i] ? worksheet['P' + i].v : 0) + ", 0) ";
                        i++;
                    
                    }
                }
                console.log('Analiza zdania WEEK ' + (weeknumber < 10 ? '0' : '') + weeknumber + ' (' + day.format('YYYY.MM.DD') + ').xlsm ILD=' + (total ? ((ok / total) * 100).toFixed(2) : 0) + ', ILD_CURE:' + (totalCure ? ((okCure / totalCure) * 100).toFixed(2) : 0));


              //WSTAWIANIE DANYCH DO BAZY - ILD  
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
                        console.log('WSTAWIONO WSKAZNIK ILD');

                    });


                  
            //PLAN Z ZAK£ADKI PLAN_ILD
                var insertIldString2 = 'INSERT ALL ';
                worksheet = workbook.Sheets["PLAN_ILD"];
                i = 6;
                if (worksheet)
                {
                    var wd11 = {};
                    while (worksheet['P' + (i+1)]) {

                        if (worksheet['P' + (i+1)].v == 'PL1') {
                            var j = 20;
                            var sap = worksheet[XLSX.utils.encode_cell({ r: i, c: 17 })].v;
                            wd11[sap] = {};
                            wd11[sap]['salecode'] = sap;
                            wd11[sap]['rozmiar'] = worksheet[XLSX.utils.encode_cell({ r: i, c: 18 })].v;
                            wd11[sap]['weeks'] = {};
                            while (worksheet[XLSX.utils.encode_cell({ r: 5, c: j })]) {

                                if (!wd11[sap]['weeks'][moment.unix((worksheet[XLSX.utils.encode_cell({ r: 5, c: j })].v - 25569) * 24 * 60 * 60).isoWeek()])
                                    wd11[sap]['weeks'][moment.unix((worksheet[XLSX.utils.encode_cell({ r: 5, c: j })].v - 25569) * 24 * 60 * 60).isoWeek()] = 0;

                                if (worksheet[XLSX.utils.encode_cell({ r: i, c: j })])
                                    wd11[sap]['weeks'][moment.unix((worksheet[XLSX.utils.encode_cell({ r: 5, c: j })].v - 25569) * 24 * 60 * 60).isoWeek()] += worksheet[XLSX.utils.encode_cell({ r: i, c: j })].v;
                                else
                                    wd11[sap]['weeks'][moment.unix((worksheet[XLSX.utils.encode_cell({ r: 5, c: j })].v - 25569) * 24 * 60 * 60).isoWeek()] += 0;
                                
                                j++;
                            }
                        }
                        i++;
                    }

                    for (var l = 0; l < Object.keys(wd11).length; l++) {
                        for (var k = 0; k < Object.keys(wd11[Object.keys(wd11)[l]]['weeks']).length; k++) {
                            insertIldString += "into ff_ild_plan(day, description,sapcode,plan_week,wd11,week) values (trunc(sysdate,'ddd'),'" + wd11[Object.keys(wd11)[l]]['rozmiar']+"'," + wd11[Object.keys(wd11)[l]]['salecode'] + "," + (wd11[Object.keys(wd11)[l]]['weeks'][Object.keys(wd11[Object.keys(wd11)[l]]['weeks'])[k]])+",1," + Object.keys(wd11[Object.keys(wd11)[l]]['weeks'])[k]+")";
                        }
                    }
                }
                insertIldString += ' SELECT * FROM DUAL';


                ////USUWANIE REKORDOW Z BAZY - ILD  PLAN
                await connection.execute('DELETE FROM FF_ILD_PLAN', {}, {},
                    function (err, result) {
                        if (err) {
                            console.error(err);
                            return;
                        }
                        console.log('USUNIÊTO DANE STARE');
                    });

                //   WSTAWIANIE DANYCH DO BAZY - ILD  PLAN
                await connection.execute(insertIldString,
                    {}, //WIAZANIE ZMIENNYCH

                    {
                        //  resultSet: true maxRows: 1000000
                        autoCommit: true
                    },
                    function (err, result) {
                        if (err) {
                            console.error(err);
                            return;
                        }
                        console.log('WSTAWIONO PLAN ILD');
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
    console.info('Execution time in node JS: %d second', end / 1000);

}

ildPut();



//fs.readdir('\\\\172.22.6.130\\pf$\\PW4 Wulkanizacja\\dzie\u0144\\ANALIZY\\', function (err, items) {
//    console.log(items);

//    for (var i = 0; i < items.length; i++) {
//        console.log(items[i]);
//    }
//});



