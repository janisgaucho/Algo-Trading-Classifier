// --- CONFIGURATION ---
const LIGNE_TITRES = 5; // Titres sur la ligne 5

// --- FONCTIONS UTILITAIRES ---

function cleanNum(str) {
    if (!str) return 0.0;
    let cleaned = str.replace(/[^\d.-]/g, '');
    return parseFloat(cleaned) || 0.0;
}

function getVal(doc, label) {
    const tds = Array.from(doc.querySelectorAll('td'));
    const target = tds.find(td => td.textContent.toLowerCase().includes(label.toLowerCase()));
    if (target && target.nextElementSibling) {
        return target.nextElementSibling.textContent.trim();
    }
    return "0";
}

function getInput(text, param) {
    const regex = new RegExp(`${param}\\s*=\\s*([\\d.]+)`, 'i');
    const match = text.match(regex);
    return match ? parseFloat(match[1]) : 0.0;
}

// --- LE "NETTOYEUR" (Crucial pour la fusion) ---
// Transforme n'importe quelle cellule (Texte riche, Nombre, Objet) en texte simple propre
function getCleanValue(cell) {
    if (!cell || cell.value === null || cell.value === undefined) return "";
    
    // Si c'est un objet (Rich Text), on extrait juste le texte
    if (typeof cell.value === 'object') {
        if (cell.value.text) return String(cell.value.text);
        if (cell.value.richText) return cell.value.richText.map(t => t.text).join('');
        if (cell.value.result) return String(cell.value.result); // Formule
    }
    
    // Sinon c'est une valeur simple
    return String(cell.value);
}

// --- LOGIQUE PRINCIPALE ---

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
        const workbook = new ExcelJS.Workbook();
        const arrayBuffer = await readFileAsArrayBuffer(excelInput.files[0]);
        await workbook.xlsx.load(arrayBuffer);
        const worksheet = workbook.worksheets[0];

        // 1. MAPPING DES COLONNES (Ligne 5)
        const mapping = {};
        const headerRow = worksheet.getRow(LIGNE_TITRES);
        
        headerRow.eachCell((cell, colNumber) => {
            const val = getCleanValue(cell).toUpperCase().replace(/[\s_\-\/']/g, '');
            if (val) mapping[val] = colNumber;
        });

        // 2. TROUVER OÙ ÉCRIRE (Première ligne vide après Ligne 5)
        let colRef = 4; // Colonne par défaut
        for (const [key, val] of Object.entries(mapping)) {
            if (key.includes("ACTIF") || key.includes("SYMBOL")) { colRef = val; break; }
        }

        let currentRow = LIGNE_TITRES + 1;
        while (true) {
            const cell = worksheet.getCell(currentRow, colRef);
            // On continue tant qu'il y a quelque chose
            if (getCleanValue(cell) !== "" || cell.isMerged) {
                currentRow++;
            } else {
                break;
            }
        }
        
        const startWritingRow = currentRow;
        let sessionData = [];

        // 3. ÉCRITURE DES DONNÉES (Sans fusionner pour l'instant)
        const files = Array.from(htmlInput.files);
        
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            statusDiv.innerHTML = `⏳ Traitement ${i + 1}/${files.length} : ${file.name}...`;

            const htmlText = await readFileAsText(file);
            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlText, 'text/html');
            const textContent = doc.body.textContent;

            // -- Extraction --
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

            // Écriture brute
            for (const [key, val] of Object.entries(DATA)) {
                const keyClean = key.toUpperCase().replace(/[\s_\-\/']/g, '');
                let colIndex = mapping[keyClean];
                
                if (!colIndex) {
                    for (const [header, idx] of Object.entries(mapping)) {
                        if (header.includes(keyClean) || keyClean.includes(header)) {
                             if (keyClean.includes("POINTS") && !header.includes("POINTS")) continue;
                             if (keyClean.includes("%") && !header.includes("%")) continue;
                             colIndex = idx; break;
                        }
                    }
                }
                if (colIndex) {
                    const cell = worksheet.getCell(currentRow, colIndex);
                    cell.value = val; // On écrit la valeur (texte ou nombre)
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                }
            }
            
            sessionData.push(DATA);
            currentRow++; 
        }

        // 4. FUSION FINALE
statusDiv.innerHTML = "🔄 Fusion des cellules identiques...";
const endWritingRow = currentRow - 1;

// On s'assure qu'on ne fusionne pas les titres (Ligne 5)
const debutFusion = LIGNE_TITRES + 1; 

if (endWritingRow > debutFusion) {
    // On boucle sur toutes les colonnes détectées dans le mapping
    Object.values(mapping).forEach(colIndex => {
        if (colIndex) {
            fusionnerColonne(worksheet, colIndex, debutFusion, endWritingRow);
        }
    });
}

        // 5. SAUVEGARDE
        statusDiv.innerHTML = "💾 Génération...";
        const buffer = await workbook.xlsx.writeBuffer();
        const finalBlob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const finalFileName = `Resultats_Complets_${Date.now()}.xlsx`;

        let cloudLink = null;
        if (window.uploadToFirebase) {
            try {
                const fileToUpload = new File([finalBlob], finalFileName, { type: finalBlob.type });
                cloudLink = await window.uploadToFirebase(fileToUpload);
            } catch (e) { console.error("Erreur Cloud", e); }
        }

        sessionData.forEach(dataItem => {
            dataItem.CLOUD_URL = cloudLink;
            saveToHistory(dataItem);
        });

        saveAs(finalBlob, finalFileName);
        statusDiv.innerHTML = `✅ Terminé ! ${files.length} fichiers traités.`;
        statusDiv.className = "status success";

    } catch (error) {
        console.error(error);
        statusDiv.innerHTML = `❌ Erreur : ${error.message}`;
        statusDiv.className = "status error";
    }
}

// FONCTION DE FUSION QUI MARCHE (Compare des Strings propres)
function fusionnerColonne(ws, colIndex, startRow, endRow) {
    let groupeDebut = startRow;

    for (let r = startRow + 1; r <= endRow + 1; r++) {
        // On récupère les valeurs en texte brut pour la comparaison
        let valRef = getCleanValue(ws.getCell(groupeDebut, colIndex)).trim();
        let valCurr = (r <= endRow) ? getCleanValue(ws.getCell(r, colIndex)).trim() : "###_FIN_###";

        // Si la valeur change, on fusionne le bloc précédent
        if (valRef !== valCurr) {
            // On ne fusionne que s'il y a au moins 2 lignes identiques
            if (r - 1 > groupeDebut) {
                try {
                    ws.mergeCells(groupeDebut, colIndex, r - 1, colIndex);
                    
                    // Style pour centrer le texte dans la cellule fusionnée
                    ws.getCell(groupeDebut, colIndex).alignment = { 
                        vertical: 'middle', 
                        horizontal: 'center',
                        wrapText: true 
                    };
                } catch (e) {
                    console.error(`Erreur fusion Col:${colIndex} Lignes:${groupeDebut}-${r-1}`, e);
                }
            }
            groupeDebut = r;
        }
    }
}

// --- LECTURE FICHIERS ---
function readFileAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsText(file, "UTF-16");
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

// --- HISTORIQUE & INTERFACE ---

function saveToHistory(dataObj) {
    let history = JSON.parse(localStorage.getItem('tradingHistory')) || [];
    dataObj.TIMESTAMP = new Date().getTime(); 
    history.unshift(dataObj);
    localStorage.setItem('tradingHistory', JSON.stringify(history));
    renderHistory();
}

function renderHistory() {
    let history = JSON.parse(localStorage.getItem('tradingHistory')) || [];
    const tbody = document.getElementById('history-body');
    if (!tbody) return;
    tbody.innerHTML = ''; 

    history.forEach((item, index) => {
        const dateObj = item.TIMESTAMP ? new Date(item.TIMESTAMP) : new Date();
        
        // --- MODIFICATION ICI ---
        const dateStr = dateObj.toLocaleDateString(); // Ex: 25/10/2023
        const hourStr = dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }); // Ex: 14:30
        // -------------------------

        // Couleur dynamique du profit
        let profitColor = String(item.NETPROFIT).includes('-') ? '#ff4d4d' : '#00C897';
        // Formatage du Profit Factor
        let pfValue = item.FACTOR || '0.00';

        let row = `
            <tr onclick="window.location.href='report-details.html?id=${index}'" 
                style="border-bottom: 1px solid #eee; cursor: pointer; transition: 0.2s;"
                onmouseover="this.style.backgroundColor='#f4f7f9'" 
                onmouseout="this.style.backgroundColor='transparent'">
                <td style="padding:15px; line-height: 1.4;">
                    ${dateStr}<br>
                    <small style="color: #888; font-weight: 400;">${hourStr}</small>
                </td>
                <td style="padding:15px; font-weight:600;">${item.ACTIF || 'N/A'}</td>
                <td style="padding:15px; color: ${profitColor}; font-weight:bold;">${item.NETPROFIT}</td>
                <td style="padding:15px; font-weight:500;">${pfValue}</td>
                <td style="padding:15px; text-align:right; color: #8A2BE2;">
                    <i class="fa-solid fa-arrow-right-long"></i>
                </td>
            </tr>`;
        tbody.innerHTML += row;
    });
}

function effacerHistorique() {
    if(confirm("Veux-tu vraiment tout effacer ?")) {
        localStorage.removeItem('tradingHistory');
        renderHistory();
    }
}
document.addEventListener('DOMContentLoaded', renderHistory);
function updateFileName(input) {
    const nameDisplay = document.getElementById(input.id === 'excelInput' ? 'excelName' : 'htmlName');
    if (input.files.length > 0) {
        nameDisplay.textContent = input.files.length === 1 ? "✅ " + input.files[0].name : "✅ " + input.files.length + " fichiers";
    } else {
        nameDisplay.textContent = "";
    }
}

// --- IA (AVEC VOTRE CLÉ) ---
function openAIAnalysis(index) {
    const history = JSON.parse(localStorage.getItem('tradingHistory')) || [];
    const reportData = history[index]; 
    if (!reportData) return;
    const modal = document.getElementById('aiModal');
    const modalBody = document.getElementById('aiModalBody');
    if (modal) {
        modal.style.display = 'flex';
        modalBody.innerHTML = `<div style="text-align:center; padding: 20px;"><i class="fa-solid fa-circle-notch fa-spin" style="font-size: 30px; color: #8A2BE2;"></i><p>Analyse IA en cours...</p></div>`;
        callGeminiAPI(reportData);
    }
}
function closeAIModal() { document.getElementById('aiModal').style.display = 'none'; }
window.onclick = function(event) { if (event.target == document.getElementById('aiModal')) closeAIModal(); }

async function callGeminiAPI(data) {
    let API_KEY = "AIzaSyBgQmEkzqrilGeD9WTLkPmkf8Ru_oiz6pw"; 
    API_KEY = API_KEY.trim();
    const modalBody = document.getElementById('aiModalBody');
    let modelName = "models/gemini-1.5-flash"; 
    try {
        const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${API_KEY}`;
        const listResp = await fetch(listUrl);
        const listData = await listResp.json();
        if (listData.models) {
            const valid = listData.models.filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"));
            const best = valid.find(m => m.name.includes("flash")) || valid.find(m => m.name.includes("pro")) || valid[0];
            if (best) modelName = best.name;
        }
    } catch (e) { console.warn("Défaut IA", e); }

    const prompt = `Agis comme un Risk Manager. Analyse ce backtest : Actif ${data.ACTIF}, Profit ${data.NETPROFIT}, Drawdown ${data.MAXDD}, Trades ${data.TRADES}, Winrate ${data.WINRATE}, PF ${data.FACTOR}. Réponds en 4 points concis (Verdict, Force, Danger, Note/10) avec balises <b>.`;

    try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${API_KEY}`, {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
        });
        const result = await response.json();
        if (result.candidates) {
            modalBody.innerHTML = `<div style="text-align:left; font-size:15px; line-height:1.6;">${result.candidates[0].content.parts[0].text.replace(/\*\*(.*?)\*\*/g, '<b>$1</b>').replace(/\n/g, '<br>')}</div>`;
        } else throw new Error("Réponse vide");
    } catch (error) {
        modalBody.innerHTML = `<p style='color:red; text-align:center;'>❌ Erreur IA: ${error.message}</p>`;
    }
}

// Fonction de secours si la première URL échoue
async function callGeminiAPI_Details(data, container) {
    const API_KEY = "AIzaSyBgQmEkzqrilGeD9WTLkPmkf8Ru_oiz6pw";
    let modelName = "models/gemini-1.5-flash"; // Valeur par défaut

    try {
        // REPRISE DE VOTRE LOGIQUE ORIGINALE : Lister les modèles pour trouver le bon
        const listResp = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${API_KEY}`);
        const listData = await listResp.json();
        
        if (listData.models) {
            const valid = listData.models.filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"));
            // On cherche flash, sinon pro, sinon le premier dispo
            const best = valid.find(m => m.name.includes("flash")) || valid.find(m => m.name.includes("pro")) || valid[0];
            if (best) modelName = best.name;
        }
    } catch (e) { 
        console.warn("Échec de la détection du modèle, utilisation du défaut", e); 
    }

    const prompt = `Agis comme un Risk Manager. Analyse ce backtest : 
    Actif ${data.ACTIF}, Profit ${data.NETPROFIT}, Drawdown ${data.MAXDD}, 
    Trades ${data.TRADES}, Winrate ${data.WINRATE}, PF ${data.FACTOR}. 
    Réponds en 4 points concis (Verdict, Force, Danger, Note/10) avec balises <b>.`;

    try {
        // L'URL utilise maintenant modelName détecté dynamiquement
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${API_KEY}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
        });

        const result = await response.json();

        if (result.candidates && result.candidates[0].content.parts[0].text) {
            let output = result.candidates[0].content.parts[0].text;
            // Conversion du format Markdown en HTML simple
            container.innerHTML = output
                .replace(/\*\*(.*?)\*\*/g, '<b>$1</b>')
                .replace(/\n/g, '<br>');
        } else {
            container.innerHTML = "L'IA n'a pas renvoyé de réponse valide.";
        }
    } catch (error) {
        console.error("Erreur Gemini:", error);
        container.innerHTML = "Erreur lors de la génération de l'analyse.";
    }
}