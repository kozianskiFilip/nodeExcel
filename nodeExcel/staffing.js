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
moment.locale('pl-PL');

async function ildPut() {
    var weeknumber;

    //TRYB_TEST
    var startDay = moment('2019-02-01');

    //TRYB NORMAL
    startDay = moment().subtract(1, 'day');
    var endDay = moment();

    console.log(startDay.format('MMMM').replace('ń','\u0144')); //zamiana znaku polskiego na charcode

    //TESTOWANIE
    // console.log(endDay.diff(startDay, 'minutes'));
    // return;


    const connection = await oracledb.getConnection({
        user: "minimes_ff_wbr",
        password: "Baza0racl3appl1cs",
        connectString: "172.22.8.47/ORA"
    });

    str = '\\\\172.22.6.130\\staffing$\\Monitoring Zatrudnienia\\___ManMinutes\\2019\\' + startDay.format("M") + '_' + startDay.format("MMMM") + '_' + startDay.format("YYYY") + '\\PK Staffing ManMin Calculation ' + startDay.format("MMMM") + '.xlsx';

    try {
        if (fs.existsSync(str)) {
            //file exists
            var workbook = XLSX.readFile(str);

            for (var day = startDay; day < endDay; day.add(1, 'day')) {
                //day = moment("02-" + d + "-2019", "MM-DD-YYYY");

                if (workbook) {
                    var worksheet = workbook.Sheets["" + day.format('D') + ""];
                    var brA = { "O": 0, "C": 0, "T": 0, "ER": 0, "I": 0, "N": 0, "S": 0, "UZ": 0, "U": 0 };// TABELA OBECNOŚCI/ABSENCJI DLA BR
                    var brB = { "O": 0, "C": 0, "T": 0, "ER": 0, "I": 0, "N": 0, "S": 0, "UZ": 0, "U": 0 };// TABELA OBECNOŚCI/ABSENCJI DLA BR
                    var brC = { "O": 0, "C": 0, "T": 0, "ER": 0, "I": 0, "N": 0, "S": 0, "UZ": 0, "U": 0 };// TABELA OBECNOŚCI/ABSENCJI DLA BR
                    var brD = { "O": 0, "C": 0, "T": 0, "ER": 0, "I": 0, "N": 0, "S": 0, "UZ": 0, "U": 0 };// TABELA OBECNOŚCI/ABSENCJI DLA BR

                    //BRYGADA A
                    if (worksheet["F70"])//OBECNOSCI
                        brA["O"] = worksheet["F70"].v;
                    if (worksheet["F67"]) //SZKOLENIA
                        brA["S"] = worksheet["F67"].v;
                    if (worksheet["F127"])
                        brA["C"] = worksheet["F127"].v;
                    if (worksheet["F128"])
                        brA["T"] = worksheet["F128"].v;
                    if (worksheet["F130"])
                        brA["ER"] = worksheet["F130"].v;
                    if (worksheet["F129"])
                        brA["I"] = worksheet["F129"].v;
                    if (worksheet["F131"])
                        brA["N"] = worksheet["F131"].v;
                    if (worksheet["F126"])
                        brA["UZ"] = worksheet["F126"].v;
                    if (worksheet["F125"])
                        brA["U"] = worksheet["F125"].v;

                    //BRYGADA B
                    if (worksheet["Q70"])//OBECNOSCI
                        brB["O"] = worksheet["Q70"].v;
                    if (worksheet["Q67"]) //SZKOLENIA
                        brB["S"] = worksheet["Q67"].v;
                    if (worksheet["Q127"])
                        brB["C"] = worksheet["Q127"].v;
                    if (worksheet["Q128"])
                        brB["T"] = worksheet["Q128"].v;
                    if (worksheet["Q130"])
                        brB["ER"] = worksheet["Q130"].v;
                    if (worksheet["Q129"])
                        brB["I"] = worksheet["Q129"].v;
                    if (worksheet["Q131"])
                        brB["N"] = worksheet["Q131"].v;
                    if (worksheet["Q126"])
                        brB["UZ"] = worksheet["Q126"].v;
                    if (worksheet["Q125"])
                        brB["U"] = worksheet["Q125"].v;

                    //BRYGADA C
                    if (worksheet["AB70"])//OBECNOSCI
                        brC["O"] = worksheet["AB70"].v;
                    if (worksheet["AB67"]) //SZKOLENIA
                        brC["S"] = worksheet["AB67"].v;
                    if (worksheet["AB127"])
                        brC["C"] = worksheet["AB127"].v;
                    if (worksheet["AB128"])
                        brC["T"] = worksheet["AB128"].v;
                    if (worksheet["AB130"])
                        brC["ER"] = worksheet["AB130"].v;
                    if (worksheet["AB129"])
                        brC["I"] = worksheet["AB129"].v;
                    if (worksheet["AB131"])
                        brC["N"] = worksheet["AB131"].v;
                    if (worksheet["AB126"])
                        brC["UZ"] = worksheet["AB126"].v;
                    if (worksheet["AB125"])
                        brC["U"] = worksheet["AB125"].v;

                    //BRYGADA D
                    if (worksheet["AM70"])//OBECNOSCI
                        brD["O"] = worksheet["AM70"].v;
                    if (worksheet["AM67"]) //SZKOLENIA
                        brD["S"] = worksheet["AM67"].v;
                    if (worksheet["AM127"])
                        brD["C"] = worksheet["AM127"].v;
                    if (worksheet["AM128"])
                        brD["T"] = worksheet["AM128"].v;
                    if (worksheet["AM130"])
                        brD["ER"] = worksheet["AM130"].v;
                    if (worksheet["AM129"])
                        brD["I"] = worksheet["AM129"].v;
                    if (worksheet["AM131"])
                        brD["N"] = worksheet["AM131"].v;
                    if (worksheet["AM126"])
                        brD["UZ"] = worksheet["AM126"].v;
                    if (worksheet["AM125"])
                        brD["U"] = worksheet["AM125"].v;

                    console.log(day.format('YY-MM-DD') + "Brygada A " + " : OBECNI: " + brA["O"] + ", URLOP: " + brA["U"] + ", CHOROBOWE: " + brA["C"] + ", TECHNICZNE: " + brA["T"]+
                        " Brygada B " + " : OBECNI: " + brB["O"] + ", URLOP: " + brB["U"] + ", CHOROBOWE: " + brB["C"] + ", TECHNICZNE: " + brB["T"]+
                        " Brygada C " + " : OBECNI: " + brC["O"] + ", URLOP: " + brC["U"] + ", CHOROBOWE: " + brC["C"] + ", TECHNICZNE: " + brC["T"]+
                        " Brygada D " + " : OBECNI: " + brD["O"] + ", URLOP: " + brD["U"] + ", CHOROBOWE: " + brD["C"] + ", TECHNICZNE: " + brD["T"]

                    );

                    //WSTAWIANIE DANYCH DLA BR. A
                    await connection.execute("update ff_data set obecni=" + brA["O"] + ", szkolenie=" + brA["S"] + ", urlop=" + brA["U"] + ", urlop_nz=" + brA["UZ"] + ", chorobowe=" + brA["C"] + ", tech=" + brA["T"] + ", er=" + brA["ER"] + ", inne=" + brA["I"] + ", nadgodziny=" + brA["N"] + " where brygada='A' and doba=to_date('" + day.format('YY-MM-DD') + "','yy-mm-dd')",
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

                    //WSTAWIANIE DANYCH DLA BR. B
                    await connection.execute("update ff_data set obecni=" + brB["O"] + ", szkolenie=" + brB["S"] + ", urlop=" + brB["U"] + ", urlop_nz=" + brB["UZ"] + ", chorobowe=" + brB["C"] + ", tech=" + brB["T"] + ", er=" + brB["ER"] + ", inne=" + brB["I"] + ", nadgodziny=" + brB["N"] + " where brygada='B' and doba=to_date('" + day.format('YY-MM-DD') + "','yy-mm-dd')",
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

                    //WSTAWIANIE DANYCH DLA BR. C
                    await connection.execute("update ff_data set obecni=" + brC["O"] + ", szkolenie=" + brC["S"] + ", urlop=" + brC["U"] + ", urlop_nz=" + brC["UZ"] + ", chorobowe=" + brC["C"] + ", tech=" + brC["T"] + ", er=" + brC["ER"] + ", inne=" + brC["I"] + ", nadgodziny=" + brC["N"] + " where brygada='C' and doba=to_date('" + day.format('YY-MM-DD') + "','yy-mm-dd')",
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

                    //WSTAWIANIE DANYCH DLA BR. D
                    await connection.execute("update ff_data set obecni=" + brD["O"] + ", szkolenie=" + brD["S"] + ", urlop=" + brD["U"] + ", urlop_nz=" + brD["UZ"] + ", chorobowe=" + brD["C"] + ", tech=" + brD["T"] + ", er=" + brD["ER"] + ", inne=" + brD["I"] + ", nadgodziny=" + brD["N"] + " where brygada='D' and doba=to_date('" + day.format('YY-MM-DD') + "','yy-mm-dd')",
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
            }
        }
        else {
            console.log('BRAK PLIKU W LOKALIZACJI: ' + str);
        }
    }
    catch (err) {
        console.error(err);
    }


    await connection.commit();
    await connection.close();
    var end = new Date() - start;
    console.info('Execution time: %dms', end);
}

ildPut();



