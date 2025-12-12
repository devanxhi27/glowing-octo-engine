function getCornerCardsEventStatsForDeclanTYPE2() {
    let newSheet1 = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/16LD3KdpEQVgrEEDtGC2GIgnOGvELrarnJcCZ6okRmtU/edit?gid=0").getSheetByName("broker/match");

    let required_clients = [
        "Mike C&C 62%",
        "Declan C&C 50%",
        "Mike C&C Mult 62%",
        "Declan C&C Mult 50%"
    ];

    let books = ['SING ALEX MAJOR', 'Grece Andreas', 'Broker Sicili Anxhelo 80%', 'SHOP BG', 'Acc Crypto', 'NEL DB-FR', 'SING Charles', 'IGC', 'SHOP ALBION', 'Fiore Italia', 'BR TR - ROME', 'ILIA Grecce', 'Bet365 BG', 'Marco', 'Baptiste', 'Stefano', 'Cico Padova', 'BR (Client)', 'Acc IT', 'Partners USA', 'Acc Asia/Exch', 'Jon Snow', 'Firence Drop', 'Israel', 'BR - Theodor RO', 'Singbet Harry', 'BR Thessaloniki', 'BR GE Canad', 'BETDEX', 'Mads', 'John', 'Lumibet', 'BR - Noah/Bettor USD', 'BR - Aus,Swiss,Ger', 'Gian', 'Shop SRB', 'BR - Valley USD', 'BR -Ari', 'BR - Brazil'];

    let book_groups = [{ "BOOKMAKER": "Shop SP-AL 80%", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "TONYS 2 80%", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "TONYS 90%", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "IGC 80%", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "IGC Tony 63.75% Eur", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "TONYS 100%", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "Tony Vietnam 45%", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "Tony Toke 63.75% Eur", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "Tony Toke 75% ALL", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "Tony Toke 63.75% ALL", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "Tony Toke 75% Eur", "GROUP 1": "SHOP ALBION" }, { "BOOKMAKER": "IGC 70%", "GROUP 1": "IGC" }, { "BOOKMAKER": "Singbet ALEX 65%", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "Singbet ALEX 60%", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "Singbet ALEX 40%", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "Probet42 Major usd", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "SING MAJOR RMB", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "SING MAJOR EARLY", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "SING MAJOR 30RMB", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "SING MAJOR 60RMB ", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "Special Alex 8a", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "SING MAJOR 40RMB ", "GROUP 1": "SING ALEX MAJOR" }, { "BOOKMAKER": "BR Hungary Shop 80%", "GROUP 1": "Grece Andreas" }, { "BOOKMAKER": "Hungary Mix shops 80%", "GROUP 1": "Grece Andreas" }, { "BOOKMAKER": "BR Montenegro Shop 80%", "GROUP 1": "Grece Andreas" }, { "BOOKMAKER": "OPAP GREECE SHOPS", "GROUP 1": "Grece Andreas" }, { "BOOKMAKER": "Andreas Greece", "GROUP 1": "Grece Andreas" }, { "BOOKMAKER": "BR Thessaloniki", "GROUP 1": "BR Thessaloniki" }, { "BOOKMAKER": "Thesaloniki OPAP  80%", "GROUP 1": "BR Thessaloniki" }, { "BOOKMAKER": "Bet365 BG 75%", "GROUP 1": "Bet365 BG" }, { "BOOKMAKER": "Bet365 75%  BG", "GROUP 1": "Bet365 BG" }, { "BOOKMAKER": "Betano 60%  BG", "GROUP 1": "Bet365 BG" }, { "BOOKMAKER": "Bet365 75%  Ricardo BG", "GROUP 1": "Bet365 BG" }, { "BOOKMAKER": "BR Alphawin BG Lev 80%", "GROUP 1": "SHOP BG" }, { "BOOKMAKER": "FF88333 80%  BG", "GROUP 1": "SHOP BG" }, { "BOOKMAKER": "BR Inbet BG Lev 80%", "GROUP 1": "SHOP BG" }, { "BOOKMAKER": "BR Sesame BG Lev 80%", "GROUP 1": "SHOP BG" }, { "BOOKMAKER": "efbet mult BG Lev 80%", "GROUP 1": "SHOP BG" }, { "BOOKMAKER": "Offline BG BET 80% ", "GROUP 1": "SHOP BG" }, { "BOOKMAKER": "BR NEL DB-FR", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "BR NEL DB-FR SBO EUR", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "BR NEL DB-FR Swisslos USD", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "BR NEL DB-FR Kambi USD", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "BR NEL DB-FR bet365 USD", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "BR NEL DB-FR SBO EUR 0.9", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "BR NEL DB-FR UNIBET USD", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "BR NEL DB-FR dayB USD", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "BR NEL DB-FR ISN USD", "GROUP 1": "NEL DB-FR" }, { "BOOKMAKER": "sports411", "GROUP 1": "Jon Snow" }, { "BOOKMAKER": "Jon Snow 3betz", "GROUP 1": "Jon Snow" }, { "BOOKMAKER": "Jon Snow ebet2", "GROUP 1": "Jon Snow" }, { "BOOKMAKER": "Jon Snow 50% USD", "GROUP 1": "Jon Snow" }, { "BOOKMAKER": "Jon Snow Xaos 50% USD", "GROUP 1": "Jon Snow" }, { "BOOKMAKER": "BR - Marco Agg", "GROUP 1": "Marco" }, { "BOOKMAKER": "BR Jacopino IN", "GROUP 1": "Marco" }, { "BOOKMAKER": "Marco Early Minori", "GROUP 1": "Marco" }, { "BOOKMAKER": "Marco Asia Bet", "GROUP 1": "Marco" }, { "BOOKMAKER": "BR - Stefano", "GROUP 1": "Stefano" }, { "BOOKMAKER": "BR - Stefano 90%", "GROUP 1": "Stefano" }, { "BOOKMAKER": "Acc Stefano 75%", "GROUP 1": "Stefano" }, { "BOOKMAKER": "BR Stefano Asia", "GROUP 1": "Stefano" }, { "BOOKMAKER": "BR TR - ROME", "GROUP 1": "BR TR - ROME" }, { "BOOKMAKER": "BR rome sing-3", "GROUP 1": "BR TR - ROME" }, { "BOOKMAKER": "Broker Sicili Anxhelo 80%", "GROUP 1": "Broker Sicili Anxhelo 80%" }, { "BOOKMAKER": "Fiore Italia", "GROUP 1": "Fiore Italia" }, { "BOOKMAKER": "Fiore Italia Asia 85% $", "GROUP 1": "Fiore Italia" }, { "BOOKMAKER": "Snai IT MI", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "Eurobet IT MIL", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "Goldbet IT MIL", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "ita tickets", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "Snai sale IT", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "Stanley sale IT", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "Planetwin365", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "Betsson", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "Betflag", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "Eplay", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "ACC EMILJAN", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "ACC Damla", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "ACC ELION", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "ACC EMILJAN 90%", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "ACC Damla 90%", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "ACC JULIAN ITALI", "GROUP 1": "Acc IT" }, { "BOOKMAKER": "BR-Betdex USD chat 2%", "GROUP 1": "BETDEX" }, { "BOOKMAKER": "BR-Betdex USD", "GROUP 1": "BETDEX" }, { "BOOKMAKER": "BR-Betdex USD online 2%", "GROUP 1": "BETDEX" }, { "BOOKMAKER": "Spain Andrea", "GROUP 1": "Spain Andrea" }, { "BOOKMAKER": "Spain Andrea 70%", "GROUP 1": "Spain Andrea" }, { "BOOKMAKER": "BR SRB DIN", "GROUP 1": "Shop SRB" }, { "BOOKMAKER": "BR SRB DIN Vra 85%", "GROUP 1": "Shop SRB" }, { "BOOKMAKER": "BR Shop SRB DIN", "GROUP 1": "Shop SRB" }, { "BOOKMAKER": "Balkan Request acc ( Antonino ) 80", "GROUP 1": "Balkan/ant" }, { "BOOKMAKER": "BR - Aus,Swiss,Ger", "GROUP 1": "BR - Aus,Swiss,Ger" }, { "BOOKMAKER": "BR - Brazil", "GROUP 1": "BR - Brazil" }, { "BOOKMAKER": "BR - Ematiq Eur", "GROUP 1": "BR - Ematiq Eur" }, { "BOOKMAKER": "BR - Noah/Bettor USD", "GROUP 1": "BR - Noah/Bettor USD" }, { "BOOKMAKER": "BR - Theodor RO", "GROUP 1": "BR - Theodor RO" }, { "BOOKMAKER": "BR - Valley USD", "GROUP 1": "BR - Valley USD" }, { "BOOKMAKER": "BR  -Ari", "GROUP 1": "BR -Ari" }, { "BOOKMAKER": "BR Alb vs BG ( singbet )", "GROUP 1": "BR Alb vs BG" }, { "BOOKMAKER": "BR Alb vs BG singbet USD", "GROUP 1": "BR Alb vs BG" }, { "BOOKMAKER": "BR Asia Eamon USD", "GROUP 1": "BR Asia Eamon USD" }, { "BOOKMAKER": "BR Betcenter Belgium", "GROUP 1": "BR Betcenter Belgium" }, { "BOOKMAKER": "BR Declan", "GROUP 1": "BR Declan" }, { "BOOKMAKER": "BR - ISR USD", "GROUP 1": "Israel" }, { "BOOKMAKER": "ISR ONLINE USD", "GROUP 1": "Israel" }, { "BOOKMAKER": "Broker Eamon USD", "GROUP 1": "John" }, { "BOOKMAKER": "BR Eamon - John", "GROUP 1": "John" }, { "BOOKMAKER": "Villiamhill", "GROUP 1": "John" }, { "BOOKMAKER": "Lumibet", "GROUP 1": "Lumibet" }, { "BOOKMAKER": "Lumibet  85%", "GROUP 1": "Lumibet" }, { "BOOKMAKER": "BR Mads 50% usd", "GROUP 1": "Mads" }, { "BOOKMAKER": "Singbet Dash", "GROUP 1": "Mads" }, { "BOOKMAKER": "Singbet Dash  60%", "GROUP 1": "Mads" }, { "BOOKMAKER": "Tennis Better - M  usd", "GROUP 1": "Mihael" }, { "BOOKMAKER": "BR Mihael eur", "GROUP 1": "Mihael" }, { "BOOKMAKER": "BR Mihael $", "GROUP 1": "Mihael" }, { "BOOKMAKER": "BR - William Minercap USA", "GROUP 1": "Minercap USA" }, { "BOOKMAKER": "BR Minercap usd", "GROUP 1": "Minercap USA" }, { "BOOKMAKER": "Partners USA", "GROUP 1": "Partners USA" }, { "BOOKMAKER": "Runshop Tennis 100%", "GROUP 1": "Runshop Tennis 100%" }, { "BOOKMAKER": "Sportwetten 80%", "GROUP 1": "Sportwetten 80%" }, { "BOOKMAKER": "BR - Tennis 222 63%", "GROUP 1": "BR - Tennis 222 63%" }, { "BOOKMAKER": "SX bet usd", "GROUP 1": "SXBet" }, { "BOOKMAKER": "ACC ITA GIAN 80%", "GROUP 1": "Gian" }, { "BOOKMAKER": "GFodds 70%", "GROUP 1": "Gian" }, { "BOOKMAKER": "BR Sale Gian 80%", "GROUP 1": "Gian" }, { "BOOKMAKER": "BR - Sale IT", "GROUP 1": "Gian" }, { "BOOKMAKER": "GFodds 80%", "GROUP 1": "Gian" }, { "BOOKMAKER": "Bookmaker", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Marathonbet", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Goldenbet", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ZodiacBet", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC Srb DIN", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Everygame US 90%", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "efbet", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "trustdice USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "vave USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "betsio USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Sbobet", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Roobet  USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Freshbet", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ISM", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC SDARJO", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "yonibet", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ERG", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ERI", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Everygame US 90% USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ERG USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Dafabet", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ELION USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ERI USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC BESO USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC MEL USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC FAT USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC FAT", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC HERA USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC MEL", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC BATI USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC TONY USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC DAN USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC India USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC Turkey USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC DARD USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC HERA", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ISM USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC Kili USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC Chile USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ARMAND USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC SOKOL USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC CANADA USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC TONYS USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC DAN", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC FINLAND USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC ARMAND", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Jose Acc 95% Eur", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC IRDI USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC KLAJDI USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC Dard USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC Bati USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC Bruno USD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC DARD", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC FAT USD 90%", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "ACC FAT USD 90% ", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Acc Austria 47.5%", "GROUP 1": "Acc Crypto" }, { "BOOKMAKER": "Punterplay 100%", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "Punterplay", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "PS3838", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "Pinnacle", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "Orbitx", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "Orbitx 95%", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "BR orbitx bet - foot", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "Exchange Rendo 50%", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "PS3838 Drinbet USD", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "Exchange LayStar", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "PS3838 USD", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "laystars 3x 3%", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "roninbet Marc mcenna 3%", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "5ball Marc Mc Enna", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "Yes2Win malayring", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "BlackBetinasia", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "Betianasia Black", "GROUP 1": "Acc Asia/Exch" }, { "BOOKMAKER": "BR Declan (Client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR Runshop (Client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR DERIVAT (client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR ES(client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR Y (client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR BETDEX (client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "Gravity (Client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR Estonia (client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR client TEMPUS", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR stefano corner (client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "BR Understated Jose (client)", "GROUP 1": "BR (Client)" }, { "BOOKMAKER": "ACC England GBP", "GROUP 1": "ACC England" }, { "BOOKMAKER": "BR ILIA gr", "GROUP 1": "ILIA Grecce" }, { "BOOKMAKER": "BR ILIA gr 95%", "GROUP 1": "ILIA Grecce" }, { "BOOKMAKER": "Singbet Baptiste 40%", "GROUP 1": "Baptiste" }, { "BOOKMAKER": "Singbet Baptiste 70%", "GROUP 1": "Baptiste" }, { "BOOKMAKER": "SBO Baptiste -1", "GROUP 1": "Baptiste" }, { "BOOKMAKER": "Br baptiste", "GROUP 1": "Baptiste" }, { "BOOKMAKER": "Singbet Charles 55%", "GROUP 1": "SING Charles" }, { "BOOKMAKER": "Singbet Charles 60%", "GROUP 1": "SING Charles" }, { "BOOKMAKER": "BR GE Canad", "GROUP 1": "BR GE Canad" }, { "BOOKMAKER": "Singbet Cico Padova 65%", "GROUP 1": "Cico Padova" }, { "BOOKMAKER": "Singbet Cico Padova 60%", "GROUP 1": "Cico Padova" }, { "BOOKMAKER": "Shop Greece", "GROUP 1": "Shop Greece" }, { "BOOKMAKER": "BR 5010Edge", "GROUP 1": "BR 5010Edge" }, { "BOOKMAKER": "BR 5010Edge Dutch", "GROUP 1": "BR 5010Edge" }, { "BOOKMAKER": "Singbet Harry 95%", "GROUP 1": "Singbet Harry" }, { "BOOKMAKER": "Firence - Drop Visar", "GROUP 1": "Firence Drop" }, { "BOOKMAKER": "Firence - Drop Lorenco", "GROUP 1": "Firence Drop" }, { "BOOKMAKER": "BR better ", "GROUP 1": "BR better" }];

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetName = spreadsheet.getName();
    // Get active sheet
    var sheet = spreadsheet.getActiveSheet();

    // Get data range and values from the data range
    var dataRange = sheet.getDataRange();
    // var values = dataRange.getValues(); // Raw values
    var displayValues = dataRange.getDisplayValues(); // Display values
    var custom_data = [];

    // Get column headers from the first row
    var headers = displayValues.shift(); // Remove and store the first row

    // Get hidden columns
    var hiddenColumns = getHiddenColumns(sheet);

    console.log("HEADERS LENGTH: " + headers.length + 2);

    // Initialize an array to store JSON objects
    var jsonData = [];

    // Iterate through each row
    for (var i = 0; i < displayValues.length; i++) {
        if (displayValues[i] && displayValues[i][3] && displayValues[i][3] !== "") {
            var rowObj = {};

            rowObj[0] = displayValues[i][60]; // Fallback to raw value if not merged

            // Iterate through each column
            for (var j = 1; j < headers.length; j++) { // Start from 1 to skip the first column
                if (hiddenColumns.indexOf(j + 1) === -1) {
                    rowObj[j] = displayValues[i][j];
                }
            }

            // Add index
            rowObj[headers.length + 2] = "INDEX_" + (i + 2);
            if (required_clients.includes(displayValues[i][26])) {
                rowObj[26] = "Declan C&C All";
                rowObj[2] = ((displayValues[i][2] == '' || displayValues[i][2] == 'L') ? 'S' : 'M');
                jsonData.push(rowObj);
            }

        }
    }

    let unique_events = [];
    let unique_event_tags = [];
    for (const itm of jsonData) {
        if (itm[4] != '' && itm[7] != '' && itm[8] != '') {
            let event_tag = new Date(itm[4]) + "__" + itm[6].replace(" ", " ") + "__" + itm[7].replace(" ", " ") + "__" + itm[8].replace(" ", " ") + "__" + ((itm[56].includes("Corn")) ? 'Corner' : itm[56].includes("Card") ? 'Card' : itm[56]);
            if (!unique_events.includes(event_tag) && (event_tag.includes("Corn") || event_tag.includes("Card") || event_tag.includes("Combination"))) {
                unique_events.push(event_tag);
            }

        }
    }
    // console.log(unique_events);
    // Sorting function
    const sorted = unique_events.sort((a, b) => {
        const [dateA, ...restA] = a.split('__');
        const [dateB, ...restB] = b.split('__');

        const dateObjA = new Date(dateA);
        const dateObjB = new Date(dateB);

        // Compare datetime first
        if (dateObjA < dateObjB) return -1;
        if (dateObjA > dateObjB) return 1;

        // If datetime is the same, compare full string alphabetically
        return a.localeCompare(b);
    });
    console.log(sorted);
    console.log("Total unique events: " + sorted.length);

    let groupedByTicketID = {};
    jsonData.forEach(obj => {
        const tag = obj[60];
        if (!groupedByTicketID[tag]) {
            groupedByTicketID[tag] = [];
        }
        groupedByTicketID[tag].push(obj);
    });
    // console.log(groupedByTicketID);
    let output_items = [];

    // Helper function to convert stake string to float, defaulting to 0
    const getStakeValue = (stake) => {
        return (stake != '' && stake != '0.00' && stake != '0') ? parseFloat(stake) : 0;
    }

    for (const event of sorted) {
        let eventIDs = [];
        let eventRecievedDate = [];
        let eventBookNamesOg = []; // New array to store original bookmaker names
        let eventStake = 0;
        let eventStakeConf = 0;
        let eventStakeOff = 0;
        let eventPL = 0;

        let eventSingleStake = 0;
        let eventSingleStakeConf = 0;
        let eventSingleStakeOff = 0;
        let eventSinglePL = 0;

        let eventMultipleStake = 0;
        let eventMultipleStakeConf = 0;
        let eventMultipleStakeOff = 0;
        let eventMultiplePL = 0;

        // Initialize bookmaker stakes object: { bookName: { single: 0, multiple: 0, total: 0 } }
        let bookmakerStakes = {};
        for (const book of books) {
            bookmakerStakes[book] = { single: 0, multiple: 0, total: 0 };
        }


        for (const [ID, itms] of Object.entries(groupedByTicketID)) {
            let hasEvent = false;
            for (const itm of itms) {
                let eventTag = new Date(itm[4]) + "__" + itm[6].replace(" ", " ") + "__" + itm[7].replace(" ", " ") + "__" + itm[8].replace(" ", " ") + "__" + ((itm[56].includes("Corn")) ? 'Corner' : itm[56].includes("Card") ? 'Card' : itm[56]);

                if (eventTag == event) {
                    hasEvent = true;
                }
            }

            if (hasEvent == true) {


                if (!eventIDs.includes(ID)) {
                    eventIDs.push(ID);
                    for (const itm of itms) {
                        const stakeValue = getStakeValue(itm[14]);
                        const plValue = getStakeValue(itm[16]);
                        const confirmedStakeValue = getStakeValue(itm[28]);
                        const offeredStakeValue = getStakeValue(itm[22]);
                        const bookNameOg = itm[18]; // Column 18 contains the bookNameOg
                        const isSingle = itm[2] == 'S';
                        const isMultiple = itm[2] == 'M';
                        const bookName = ((book_groups.filter(item => item["BOOKMAKER"] == bookNameOg).length > 0) ? book_groups.filter(item => item["BOOKMAKER"] == bookNameOg)[0]["GROUP 1"] : bookNameOg);

                        // **Collect original bookmaker name**
                        if (!eventBookNamesOg.includes(bookNameOg)) {
                            eventBookNamesOg.push(bookNameOg);
                        }

                        eventStake = eventStake + stakeValue;
                        eventPL = eventPL + plValue;

                        eventStakeConf = eventStakeConf + confirmedStakeValue;


                        eventStakeOff = eventStakeOff + offeredStakeValue;

                        if (isSingle) {
                            eventSingleStake = eventSingleStake + stakeValue;
                            eventSinglePL = eventSinglePL + plValue;

                            eventSingleStakeConf = eventSingleStakeConf + confirmedStakeValue;


                            eventSingleStakeOff = eventSingleStakeOff + offeredStakeValue;
                        }


                        if (isMultiple) {
                            eventMultipleStake = eventMultipleStake + stakeValue;
                            eventMultiplePL = eventMultiplePL + plValue;

                            eventMultipleStakeConf = eventMultipleStakeConf + confirmedStakeValue;


                            eventMultipleStakeOff = eventMultipleStakeOff + offeredStakeValue;
                        }

                        // New logic for bookmaker stakes
                        // Only update bookmaker stakes if the bookName is one of the books in the 'books' array
                        if (bookmakerStakes[bookName]) {
                            if (isSingle) {
                                bookmakerStakes[bookName].single += stakeValue;
                            }
                            if (isMultiple) {
                                bookmakerStakes[bookName].multiple += stakeValue;
                            }
                            bookmakerStakes[bookName].total += stakeValue;
                        }


                        if (!eventRecievedDate.includes(itm[3])) {
                            eventRecievedDate.push(itm[3]);
                        }
                    }
                }


            }
        }

        // Prepare bookmaker stakes array for output
        let bookStakesOutput = [];
        for (const book of books) {
            const stakes = bookmakerStakes[book];
            // Order: bookname single stake, bookname multiple stake, bookname total stake
            bookStakesOutput.push(stakes.single || 0, stakes.multiple || 0, stakes.total || 0);
        }

        // **Push the updated object with bookmaker stakes and BookNamesOg**
        output_items.push({
            event: event,
            eventIDs: eventIDs,
            eventRecievedDate: eventRecievedDate,
            eventBookNamesOg: eventBookNamesOg, // New: Array of original bookmaker names
            eventStake: eventStake,
            eventPL: eventPL,
            eventStakeConf: eventStakeConf,
            eventStakeOff: eventStakeOff,
            eventSingleStake: eventSingleStake,
            eventSinglePL: eventSinglePL,
            eventSingleStakeConf: eventSingleStakeConf,
            eventSingleStakeOff: eventSingleStakeOff,
            eventMultipleStake: eventMultipleStake,
            eventMultiplePL: eventMultiplePL,
            eventMultipleStakeConf: eventMultipleStakeConf,
            eventMultipleStakeOff: eventMultipleStakeOff,
            bookStakesOutput: bookStakesOutput // The new array of bookmaker stakes
        });

        // console.log(ttest);
    }

    var spreadsheetName = spreadsheet.getName();
    console.log(output_items[20]);
    let output_data = [];
    for (const output_item of output_items) {
        let eventTime = output_item.event.split("__")[0];
        let eventLeague = output_item.event.split("__")[1];
        let eventHome = output_item.event.split("__")[2];
        let eventAway = output_item.event.split("__")[3];
        let eventLine = output_item.event.split("__")[4];

        let eventTitle = eventHome + " vs " + eventAway;

        // **Update output_data push to include bookStakesOutput and BookNamesOg**
        output_data.push([
            spreadsheetName,
            eventLeague,
            eventTime,
            eventTitle,
            eventLine,
            output_item.eventStake,
            //    output_item.eventStakeConf,
            //    output_item.eventStakeOff,
            output_item.eventSingleStake,
            //    output_item.eventSingleStakeConf,
            //    output_item.eventSingleStakeOff,
            output_item.eventMultipleStake,
            //    output_item.eventMultipleStakeConf,
            //    output_item.eventMultipleStakeOff,
            output_item.eventPL,
            output_item.eventSinglePL,
            output_item.eventMultiplePL,
            output_item.eventIDs.join(", "),
            output_item.eventRecievedDate.join(", "),
            ...output_item.bookStakesOutput, // Add the new bookmaker stake values
            output_item.eventBookNamesOg.join(", ") // New: Original Bookmaker Names (last column)
        ])
    }
    console.log(output_data[20]);
    // console.log(ttest);


    console.log("Total ROWS: " + output_data.length);

    newSheet1.getRange(2223, 1, output_data.length, output_data[0].length).setValues(output_data);
}