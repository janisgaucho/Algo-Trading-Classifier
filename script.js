const LIGNE_TITRES = 5;

// --- FONCTIONS UTILITAIRES --- //

function cleanNum(str) {
    if (!str) return 0.0;
    // Enlève tout sauf chiffres, point et moins
    let cleaned = str.replace(/[^\d.-]/g, '');
    return parseFloat(cleaned) || 0.0;
}

function getVal(doc, label) {
    // Cherche un <td> qui contient le label
    const tds = Array.from(doc.querySelectorAll('td'));
    const target = tds.find(td => td.textContent.toLowerCase().includes(label.toLowerCase()));
    if (target && target.nextElementSibling) {
        return target.nextElementSibling.textContent.trim();
    }
    return "0";
}

function getInput(text, param) {
    // Regex équivalent au python
    const regex = new RegExp(`${param}\\s*=\\s*([\\d.]+)`, 'i');
    const match = text.match(regex);
    return match ? parseFloat(match[1]) : 0.0;
}

// --- LOGIQUE PRINCIPALE --- //

async function processFiles() {
    const excelInput = document.getElementById('excelInput');
    const htmlInput = document.getElementById('htmlInput');
    const statusDiv = document.getElementById('status');

    if (!excelInput.files[0] || htmlInput.files.length === 0) {
        statusDiv.innerHTML = "⚠️ Veuillez sélectionner le fichier Excel et au moins un rapport HTML.";
        statusDiv.className = "status error";
        return;
    }

    statusDiv.innerHTML = "⏳ Chargement du fichier Excel...";
    statusDiv.className = "status";

    try {
        // --- 1. PRÉPARATION ---
        const workbook = new ExcelJS.Workbook();
        const arrayBuffer = await readFileAsArrayBuffer(excelInput.files[0]);
        await workbook.xlsx.load(arrayBuffer);
        const worksheet = workbook.worksheets[0];

        // Mapping des colonnes (Code habituel)
        const mapping = {};
        const headerRow = worksheet.getRow(LIGNE_TITRES);
        headerRow.eachCell((cell, colNumber) => {
            if (cell.value) {
                let cleanHeader = String(cell.value).toUpperCase().replace(/[\s_\-\/']/g, '');
                mapping[cleanHeader] = colNumber;
            }
        });

        let colRef = 4; // Par défaut
        for (const [key, val] of Object.entries(mapping)) {
            if (key.includes("ACTIF") || key.includes("SYMBOL")) { colRef = val; break; }
        }

        let currentRow = LIGNE_TITRES + 1;
        while (true) {
            const cell = worksheet.getCell(currentRow, colRef);
            if ((cell.value !== null && cell.value !== "") || cell.isMerged) currentRow++;
            else break;
        }

        // Variable temporaire pour stocker les infos de la session avant sauvegarde
        let sessionData = [];

        // --- 2. BOUCLE DE TRAITEMENT (Remplissage) ---
        const files = Array.from(htmlInput.files);
        
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            statusDiv.innerHTML = `⏳ Traitement du fichier ${i + 1} sur ${files.length} : ${file.name}...`;

            const htmlText = await readFileAsText(file);
            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlText, 'text/html');
            const textContent = doc.body.textContent;

            // --- EXTRACTION (Ton code habituel) ---
            const expertName = getVal(doc, "Expert");
            const symbol = getVal(doc, "Symbol");
            const rawPeriod = getVal(doc, "Period");
            const trades = cleanNum(getVal(doc, "Total Trades"));
            const profit = cleanNum(getVal(doc, "Total Net Profit"));
            
            let initDep = cleanNum(getVal(doc, "Initial Deposit"));
            if (initDep === 0) initDep = 100000.0;

            const vol = getInput(textContent, "Entry_Amount");
            const sl = getInput(textContent, "Stop_Loss");
            const tp = getInput(textContent, "Take_Profit");
            const maxOpenPos = getInput(textContent, "Max_OpenPos");

            const maxDDStr = getVal(doc, "Equity Drawdown Maximal");
            let maxDDPct = 0.0;
            const maxDDMatch = maxDDStr.match(/\(([\d.]+)%\)/);
            if (maxDDMatch) maxDDPct = parseFloat(maxDDMatch[1]);
            const maxDDDisplay = `${maxDDPct}%`;

            const relDDStr = getVal(doc, "Relative Drawdown");
            let relDDPct = 0.0;
            if (relDDStr && relDDStr.includes("%")) {
                const relMatch = relDDStr.match(/([\d.]+)%/);
                if (relMatch) relDDPct = parseFloat(relMatch[1]);
            }
            if (relDDPct === 0) relDDPct = maxDDPct;

            const maxLossNb = getVal(doc, "Maximum consecutive losses").split('(')[0].trim();
            const dateRegex = /\d{4}\.\d{2}\.\d{2}/g;
            const dates = rawPeriod.match(dateRegex) || [];
            const dStart = dates.length > 0 ? dates[0] : "";
            const dEnd = dates.length > 0 ? dates[dates.length - 1] : "";
            let duree = 0;
            if (dStart && dEnd) duree = parseInt(dEnd.substring(0, 4)) - parseInt(dStart.substring(0, 4));
            const timeframe = rawPeriod.split('(')[0].trim();

            const gainsVal = (profit * 100) / initDep;
            const gainsDisplay = `${gainsVal.toFixed(2)}%`;
            const gainMensuelCash = profit / 60;
            const moyMensuelPct = parseFloat(((gainMensuelCash / initDep) * 100).toFixed(2));
            let retDD = 0;
            if (relDDPct !== 0) retDD = parseFloat((moyMensuelPct / relDDPct).toFixed(2));
            const mTrades = parseFloat((trades / 60).toFixed(2));
            const slPct = parseFloat(((sl * vol) / initDep).toFixed(5));
            const tpPct = parseFloat(((tp * vol) / initDep).toFixed(5));
            let winrate = "";
            const profitTradesStr = getVal(doc, "Profit Trades");
            const winMatch = profitTradesStr.match(/\(([\d.]+%)\)/);
            if (winMatch) winrate = winMatch[1];
            const profitFactor = getVal(doc, "Profit Factor");
            const sharpe = getVal(doc, "Sharpe Ratio");

            const DATA = {
                'NOMDELEA': expertName, 'NOM': expertName, 'EXPERT': expertName,
                'ACTIF': symbol, 'INITIALBALANCE': initDep, 'INITIALDEPOSIT': initDep,
                'DEBUT': dStart, 'FIN': dEnd, 'DUREE': duree,
                'PERIOD': timeframe, 'TIMEFRAME': timeframe,
                'LOT': vol, 'VOLUME': vol, 'MAXORDERSAMETIME': maxOpenPos, 'MAXORDER': maxOpenPos,
                'TRADES': trades, 'MOYEN/MOIS': mTrades,
                'GAINS': gainsDisplay, 'NETPROFIT': gainsDisplay,
                '%MOYEN': moyMensuelPct, 'WINRATE': winrate,
                'FACTOR': profitFactor, 'SHARP': sharpe,
                'RETURN': retDD, 'MAXDD': maxDDDisplay, 'PERTES': maxLossNb,
                'SLENPOINTS': sl, 'SLEN%': slPct, 'TPENPOINTS': tp, 'TPEN%': tpPct
            };

            // Ecriture dans l'Excel en mémoire
            for (const [key, val] of Object.entries(DATA)) {
                const keyClean = key.toUpperCase().replace(/[\s_\-\/']/g, '');
                let colIndex = mapping[keyClean];
                if (!colIndex) {
                    for (const [header, idx] of Object.entries(mapping)) {
                        if (header.includes(keyClean) || keyClean.includes(header)) {
                            if (keyClean.includes("POINTS") && !header.includes("POINTS")) continue;
                            if (keyClean.includes("%") && !header.includes("%")) continue;
                            colIndex = idx;
                            break;
                        }
                    }
                }
                if (colIndex) writeAndMerge(worksheet, currentRow, colIndex, val);
            }
            
            // On ajoute les données dans notre liste temporaire
            sessionData.push(DATA);
            currentRow++;
        }

        // --- 3. GÉNÉRATION DU FICHIER FINAL ---
        statusDiv.innerHTML = "💾 Génération du fichier final...";
        const buffer = await workbook.xlsx.writeBuffer();
        
        // On crée le BLOB final (C'est le fichier Excel REMPLI)
        const finalBlob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const finalFileName = `Resultats_Complets_${Date.now()}.xlsx`;

        // --- 4. UPLOAD VERS FIREBASE (Le fichier REMPLI) ---
        let cloudLink = null;
        if (window.uploadToFirebase) {
            statusDiv.innerHTML = `☁️ Sauvegarde du fichier complet vers le Cloud...`;
            try {
                // On transforme le Blob en "Fichier" pour que Firebase comprenne
                const fileToUpload = new File([finalBlob], finalFileName, { type: finalBlob.type });
                
                // UPLOAD !
                cloudLink = await window.uploadToFirebase(fileToUpload);
            } catch (e) {
                console.error("Erreur Cloud", e);
            }
        }

        // --- 5. SAUVEGARDE HISTORIQUE ---
        // On met à jour chaque ligne de l'historique avec le LIEN DU FICHIER FINAL
        sessionData.forEach(dataItem => {
            dataItem.CLOUD_URL = cloudLink; // On attache le lien
            saveToHistory(dataItem);
        });

        // --- 6. TÉLÉCHARGEMENT LOCAL ---
        saveAs(finalBlob, finalFileName);

        statusDiv.innerHTML = `✅ Terminé ! ${files.length} analyses sauvegardées.`;
        statusDiv.className = "status success";

    } catch (error) {
        console.error(error);
        statusDiv.innerHTML = `❌ Erreur : ${error.message}`;
        statusDiv.className = "status error";
    }
}

// --- FONCTION DE FUSION (VERSION JS) ---
function writeAndMerge(ws, row, col, val) {
    const cell = ws.getCell(row, col);
    cell.value = val;

    // Centrage par défaut
    cell.alignment = { vertical: 'middle', horizontal: 'center' };

    const rowAbove = row - 1;
    if (rowAbove <= LIGNE_TITRES) return;

    const cellAbove = ws.getCell(rowAbove, col);
    
    // Comparaison (On convertit en string pour comparer)
    // On doit récupérer la valeur "maître" si c'est fusionné
    let valAbove = cellAbove.value;
    if (cellAbove.isMerged && cellAbove.master) {
        valAbove = cellAbove.master.value;
    }

    // Fonction de comparaison souple
    const isSame = (v1, v2) => {
        if (v1 == v2) return true; // Egalité simple (inclut 100 == "100")
        try {
            // Comparaison numérique
            const n1 = parseFloat(String(v1).replace(/[\s,]/g, ''));
            const n2 = parseFloat(String(v2).replace(/[\s,]/g, ''));
            if (!isNaN(n1) && !isNaN(n2) && Math.abs(n1 - n2) < 0.0001) return true;
        } catch(e) {}
        return false;
    };

    if (isSame(val, valAbove)) {
        // Logique de fusion ExcelJS
        let startRow = rowAbove;
        
        // Si celle du dessus est déjà fusionnée, on récupère le début de sa fusion
        if (cellAbove.isMerged) {
            // master renvoie la cellule en haut à gauche de la fusion
            startRow = parseInt(cellAbove.master.row); 
        }

        // On fusionne du début du bloc précédent jusqu'à la ligne actuelle
        try {
            ws.mergeCells(startRow, col, row, col);
        } catch (e) {
            console.log("Erreur fusion, on continue...");
        }
    }
}

// Helpers pour lire les fichiers
function readFileAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsText(file, "UTF-16"); // Essai UTF-16 d'abord (MT4/5)
    });
}

function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsArrayBuffer(file);
    });
}

// --- GESTION DE LA MÉMOIRE (LOCALSTORAGE) ---

// 1. Sauvegarder
function saveToHistory(dataObj) {
    let history = JSON.parse(localStorage.getItem('tradingHistory')) || [];
    
    // On ajoute un Timestamp précis pour pouvoir extraire Date et Heure séparément
    dataObj.TIMESTAMP = new Date().getTime(); 
    
    history.unshift(dataObj);
    localStorage.setItem('tradingHistory', JSON.stringify(history));
    
    renderHistory();
}

// 2. Afficher (Avec les nouvelles colonnes)
function renderHistory() {
    let history = JSON.parse(localStorage.getItem('tradingHistory')) || [];
    const tbody = document.getElementById('history-body');
    if (!tbody) return;

    tbody.innerHTML = ''; 

    history.forEach((item, index) => {
        const dateObj = item.TIMESTAMP ? new Date(item.TIMESTAMP) : new Date();
        const dateStr = dateObj.toLocaleDateString();
        const timeStr = dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        
        // Couleur profit
        let profitColor = 'black';
        let profitVal = String(item.NETPROFIT || "0");
        if(profitVal.includes('-')) profitColor = '#ff4d4d';
        else if(parseFloat(profitVal) > 0) profitColor = '#00C897';

        // --- LE BOUTON UNIQUE ---
        let actionBtn = `<span style="color:#999; font-size:12px;">Non dispo</span>`;
        
        if (item.CLOUD_URL && item.CLOUD_URL.startsWith('http')) {
            // C'est le bouton magique qui télécharge le VRAI fichier
            actionBtn = `
                <a href="${item.CLOUD_URL}" target="_blank" style="text-decoration:none;">
                        <i class="fa-solid fa-download"></i> Télécharger le rapport
                </a>`;
        }
        // ------------------------

        let row = `
            <tr style="border-bottom: 1px solid #eee;">
                <td style="padding:10px;">${dateStr} <br> <small style="color:#888">${timeStr}</small></td>
                <td style="padding:10px; font-weight:bold;">${item.ACTIF || item.SYMBOL || '-'}</td>
                <td style="padding:10px; color: ${profitColor}; font-weight:bold;">${item.NETPROFIT}</td>
                <td style="padding:10px; text-align:center;">
                    ${actionBtn}
                </td>
            </tr>
        `;
        tbody.innerHTML += row;
    });
}

// 3. Effacer l'historique
function effacerHistorique() {
    if(confirm("Veux-tu vraiment tout effacer ?")) {
        localStorage.removeItem('tradingHistory');
        renderHistory();
    }
}

// 4. Charger l'historique au démarrage de la page
document.addEventListener('DOMContentLoaded', renderHistory);

// 3. Fonction pour télécharger depuis l'historique
async function downloadFromHistory(index) {
    let history = JSON.parse(localStorage.getItem('tradingHistory')) || [];
    const data = history[index];

    if (!data) return alert("Erreur : Données introuvables.");

    try {
        // Création d'un nouveau classeur simple
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Rapport');

        // On définit les colonnes basées sur les clés de tes données
        const columns = [];
        const values = [];

        for (const [key, val] of Object.entries(data)) {
            // On ignore le timestamp interne
            if(key === 'TIMESTAMP' || key === 'DATE_ANALYSE') continue;
            
            columns.push({ header: key, key: key, width: 15 });
            values.push(val);
        }

        worksheet.columns = columns;
        
        // Ajout de la ligne de données
        worksheet.addRow(values);

        // Style rapide (En-tête gras)
        worksheet.getRow(1).font = { bold: true };

        // Génération et téléchargement
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        
        // Nom du fichier : Rapport_Symbol_Date.xlsx
        const fileName = `Rapport_${data.ACTIF || 'Unknown'}_${new Date().getTime()}.xlsx`;
        saveAs(blob, fileName);

    } catch (e) {
        console.error(e);
        alert("Erreur lors de la génération du fichier : " + e.message);
    }
}

// Fonction cosmétique pour afficher le nom des fichiers sélectionnés dans les Dropzones
function updateFileName(input) {
    const nameDisplay = document.getElementById(input.id === 'excelInput' ? 'excelName' : 'htmlName');
    if (input.files.length > 0) {
        if (input.files.length === 1) {
            nameDisplay.textContent = "✅ " + input.files[0].name;
        } else {
            nameDisplay.textContent = "✅ " + input.files.length + " fichiers sélectionnés";
        }
    } else {
        nameDisplay.textContent = "";
    }
}