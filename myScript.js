// Variables globales
let selectedFiles = [];
let repairResults = [];

// Récupération des éléments DOM
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const filesSection = document.getElementById('filesSection');
const filesList = document.getElementById('filesList');
const fileCount = document.getElementById('fileCount');
const repairBtn = document.getElementById('repairBtn');
const resultsSection = document.getElementById('resultsSection');
const resultsList = document.getElementById('resultsList');
const downloadAllBtn = document.getElementById('downloadAllBtn');

// Event Listeners
uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '#3b82f6';
    uploadArea.style.background = '#f9fafb';
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.style.borderColor = '#d1d5db';
    uploadArea.style.background = 'transparent';
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '#d1d5db';
    uploadArea.style.background = 'transparent';
    handleFiles(e.dataTransfer.files);
});

fileInput.addEventListener('change', (e) => {
    handleFiles(e.target.files);
});

// Fonction de gestion des fichiers sélectionnés
function handleFiles(files) {
    const validExtensions = ['xlsx', 'xls', 'docx', 'doc', 'pptx', 'ppt'];
    selectedFiles = Array.from(files).filter(file => {
        const ext = file.name.split('.').pop().toLowerCase();
        return validExtensions.includes(ext);
    });

    if (selectedFiles.length > 0) {
        displayFiles();
        filesSection.style.display = 'block';
        resultsSection.style.display = 'none';
    } else {
        alert('Aucun fichier Office valide détecté.\nFormats acceptés : .xlsx, .xls, .docx, .doc, .pptx, .ppt');
    }
}

// Fonction pour obtenir l'icône du fichier selon son type
function getFileIcon(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    
    if (['xlsx', 'xls'].includes(ext)) {
        return `<svg class="file-icon" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
        </svg>`;
    } else if (['docx', 'doc'].includes(ext)) {
        return `<svg class="file-icon" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
        </svg>`;
    } else if (['pptx', 'ppt'].includes(ext)) {
        return `<svg class="file-icon" viewBox="0 0 24 24" fill="none" stroke="#ea580c" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <rect x="8" y="12" width="8" height="6"></rect>
        </svg>`;
    }
    
    // Icône par défaut
    return `<svg class="file-icon" viewBox="0 0 24 24" fill="none" stroke="#6b7280" stroke-width="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
        <polyline points="14 2 14 8 20 8"></polyline>
    </svg>`;
}

// Fonction pour afficher la liste des fichiers sélectionnés
function displayFiles() {
    filesList.innerHTML = '';
    fileCount.textContent = selectedFiles.length;

    selectedFiles.forEach(file => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            ${getFileIcon(file.name)}
            <span class="file-name">${file.name}</span>
            <span class="file-size">${(file.size / 1024).toFixed(2)} KB</span>
        `;
        filesList.appendChild(fileItem);
    });
}

// Event listener du bouton de restauration
repairBtn.addEventListener('click', async () => {
    repairBtn.disabled = true;
    repairBtn.innerHTML = `
        <div class="spinner"></div>
        Restauration en cours...
    `;

    repairResults = [];
    
    for (const file of selectedFiles) {
        const result = await createBlankFile(file);
        repairResults.push(result);
    }

    displayResults();
    repairBtn.disabled = false;
    repairBtn.innerHTML = 'Restaurer à la version de base';
});

// Fonction principale de création de fichier vierge
async function createBlankFile(file) {
    try {
        const ext = file.name.split('.').pop().toLowerCase();
        let blob;

        // Utilisation d'un switch pour plus de clarté
        switch (ext) {
            case 'xlsx':
            case 'xls':
                blob = await createBlankExcel();
                break;
            case 'docx':
            case 'doc':
                blob = await createBlankWord();
                break;
            case 'pptx':
            case 'ppt':
                blob = await createBlankPowerPoint();
                break;
            default:
                throw new Error('Format de fichier non supporté');
        }

        // Vérification que le blob a bien été créé
        if (!blob || blob.size === 0) {
            throw new Error('Échec de la création du fichier');
        }

        return {
            name: file.name,
            success: true,
            blob: blob,
            message: 'Fichier restauré avec succès (version vierge créée)'
        };
    } catch (error) {
        console.error('Erreur création fichier:', error);
        return {
            name: file.name,
            success: false,
            message: 'Erreur lors de la restauration: ' + error.message
        };
    }
}

// Fonction pour créer un fichier Excel vierge
async function createBlankExcel() {
    try {
        const zip = new JSZip();
        
        // Structure minimale d'un fichier Excel
        zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`);

        zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);

        const xl = zip.folder('xl');
        xl.file('workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Feuille1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>`);

        xl.folder('_rels').file('workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);

        xl.folder('worksheets').file('sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData/>
</worksheet>`);

        const blob = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        
        return blob;
    } catch (error) {
        console.error('Erreur création Excel:', error);
        throw error;
    }
}

// Fonction pour créer un fichier Word vierge
async function createBlankWord() {
    try {
        const zip = new JSZip();
        
        zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

        zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

        zip.folder('word').file('document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p/>
</w:body>
</w:document>`);

        const blob = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        
        return blob;
    } catch (error) {
        console.error('Erreur création Word:', error);
        throw error;
    }
}

// Fonction pour créer un fichier PowerPoint vierge
async function createBlankPowerPoint() {
    try {
        const zip = new JSZip();
        
        zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`);

        zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

        const ppt = zip.folder('ppt');
        ppt.file('presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:sldIdLst>
<p:sldId id="256" r:id="rId1"/>
</p:sldIdLst>
</p:presentation>`);

        ppt.folder('_rels').file('presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`);

        ppt.folder('slides').file('slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:spTree>
<p:nvGrpSpPr>
<p:cNvPr id="1" name=""/>
<p:cNvGrpSpPr/>
<p:nvPr/>
</p:nvGrpSpPr>
<p:grpSpPr/>
</p:spTree>
</p:cSld>
</p:sld>`);

        const blob = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        });
        
        return blob;
    } catch (error) {
        console.error('Erreur création PowerPoint:', error);
        throw error;
    }
}

// Fonction pour afficher les résultats
function displayResults() {
    resultsSection.style.display = 'block';
    resultsList.innerHTML = '';

    const hasSuccess = repairResults.some(r => r.success);
    downloadAllBtn.style.display = hasSuccess ? 'flex' : 'none';

    repairResults.forEach((result, index) => {
        const resultItem = document.createElement('div');
        resultItem.className = `result-item ${result.success ? 'result-success' : 'result-error'}`;
        
        const icon = result.success 
            ? `<svg class="result-icon" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
                <polyline points="22 4 12 14.01 9 11.01"></polyline>
               </svg>`
            : `<svg class="result-icon" viewBox="0 0 24 24" fill="none" stroke="#dc2626" stroke-width="2">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="8" x2="12" y2="12"></line>
                <line x1="12" y1="16" x2="12.01" y2="16"></line>
               </svg>`;

        resultItem.innerHTML = `
            ${icon}
            <div class="result-content">
                <div class="result-name">${result.name}</div>
                <div class="result-message">${result.message}</div>
            </div>
            ${result.success ? `
                <button class="download-btn" onclick="downloadFile(${index})">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                        <polyline points="7 10 12 15 17 10"></polyline>
                        <line x1="12" y1="15" x2="12" y2="3"></line>
                    </svg>
                </button>
            ` : ''}
        `;
        resultsList.appendChild(resultItem);
    });
}

// Fonction pour télécharger un fichier individuel
function downloadFile(index) {
    try {
        const result = repairResults[index];
        
        if (!result) {
            console.error('Résultat non trouvé à l\'index:', index);
            return;
        }
        
        if (result.blob) {
            const url = URL.createObjectURL(result.blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'restaure_' + result.name;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            
            // Nettoyer après un délai
            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 100);
        } else {
            console.error('Aucun blob disponible pour le téléchargement');
            alert('Erreur : aucun fichier disponible pour le téléchargement');
        }
    } catch (error) {
        console.error('Erreur lors du téléchargement:', error);
        alert('Erreur lors du téléchargement du fichier');
    }
}

// Event listener pour télécharger tous les fichiers
downloadAllBtn.addEventListener('click', () => {
    const successfulResults = repairResults
        .map((result, index) => ({ result, originalIndex: index }))
        .filter(item => item.result.success);
    
    if (successfulResults.length === 0) {
        alert('Aucun fichier à télécharger');
        return;
    }
    
    // Télécharger avec un délai pour éviter les problèmes de navigateur
    successfulResults.forEach((item, index) => {
        setTimeout(() => {
            downloadFile(item.originalIndex);
        }, index * 300);
    });
});
