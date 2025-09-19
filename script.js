// script.js
// Конвертация Excel‑файла в формат Avito XML v3.0 в браузере

// Ссылки на элементы интерфейса
const fileInput = document.getElementById('file-input');
const convertBtn = document.getElementById('convert-btn');
const statusBox = document.getElementById('status');

// Включаем кнопку, когда файл выбран
fileInput.addEventListener('change', () => {
    convertBtn.disabled = !fileInput.files.length;
    statusBox.style.display = 'none';
    statusBox.textContent = '';
});

// Основной обработчик конвертации
convertBtn.addEventListener('click', () => {
    if (!fileInput.files.length) return;
    const file = fileInput.files[0];
    statusBox.style.display = 'block';
    statusBox.textContent = 'Чтение файла…';
    // Читаем файл как ArrayBuffer
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            // Находим лист с данными: выбираем первый лист, который содержит строку с названием категории «Готовый бизнес»
            let sheetName = workbook.SheetNames[0];
            workbook.SheetNames.forEach(name => {
                if (/Готовый\s+бизнес/i.test(name) || /Торговля/i.test(name)) {
                    sheetName = name;
                }
            });
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) {
                throw new Error('Не удалось найти нужный лист в таблице');
            }
            // Преобразуем весь лист в массив массивов (двумерный массив)
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            // Проверяем структуру: первая строка (index 0) — название категории, index 1 — названия колонок
            if (rows.length < 5) {
                throw new Error('Лист содержит недостаточно строк');
            }
            const headerRow = rows[1];
            // Сопоставляем названия колонок с индексами
            const colIndex = {};
            headerRow.forEach((title, idx) => {
                colIndex[title] = idx;
            });
            // Получаем индексы нужных колонок
            const idxMap = {
                id: colIndex['Уникальный идентификатор объявления'],
                adId: colIndex['Номер объявления на Авито'],
                manager: colIndex['Контактное лицо'],
                phone: colIndex['Номер телефона'],
                address: colIndex['Адрес'],
                title: colIndex['Название объявления'],
                description: colIndex['Описание объявления'],
                price: colIndex['Цена'],
                photos: colIndex['Ссылки на фото'],
                contactMethod: colIndex['Способ связи'],
                category: colIndex['Категория'],
                businessType: colIndex['Вид бизнеса'],
                goodsSubtype: colIndex['Вид франшизы'],
                franchiseSubtype: colIndex['Тип франшизы'],
                fee: colIndex['Паушальный взнос'],
                royalty: colIndex['Роялти'],
                royaltyType: colIndex['Тип роялти'],
                fixedRoyalty: colIndex['Фиксированное роялти'],
                percentRoyalty: colIndex['Процентное роялти'],
                payback: colIndex['Окупаемость франшизы'],
                support: colIndex['Сопровождение'],
                supportType: colIndex['Тип сопровождения'],
                audience: colIndex['Целевая аудитория'],
                status: colIndex['AvitoStatus'],
                email: colIndex['Почта'],
                dateEnd: colIndex['AvitoDateEnd'],
                video: colIndex['Ссылка на видео'],
                company: colIndex['Название компании']
            };
            // Конвертируем строки (начиная с index 4) в XML
            let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<Ads formatVersion="3" target="Avito.ru">\n';
            for (let i = 4; i < rows.length; i++) {
                const row = rows[i];
                // Пропускаем полностью пустые строки
                if (!row || row.length === 0 || !row[idxMap.id]) continue;
                xml += '  <Ad>\n';
                // Id (обязательный)
                xml += `    <Id>${escapeXml(row[idxMap.id])}</Id>\n`;
                // AdId (необязательный)
                if (row[idxMap.adId]) {
                    xml += `    <AdId>${escapeXml(row[idxMap.adId])}</AdId>\n`;
                }
                // ManagerName
                if (row[idxMap.manager]) {
                    xml += `    <ManagerName>${escapeXml(row[idxMap.manager])}</ManagerName>\n`;
                }
                // ContactPhone
                if (row[idxMap.phone]) {
                    xml += `    <ContactPhone>${escapeXml(row[idxMap.phone])}</ContactPhone>\n`;
                }
                // Address
                if (row[idxMap.address]) {
                    xml += `    <Address>${escapeXml(row[idxMap.address])}</Address>\n`;
                }
                // Category (фиксированное значение)
                xml += '    <Category>Готовый бизнес</Category>\n';
                // BusinessType
                if (row[idxMap.businessType]) {
                    xml += `    <BusinessType>${escapeXml(row[idxMap.businessType])}</BusinessType>\n`;
                }
                // GoodsSubType (Вид франшизы)
                if (row[idxMap.goodsSubtype]) {
                    xml += `    <GoodsSubType>${escapeXml(row[idxMap.goodsSubtype])}</GoodsSubType>\n`;
                }
                // FranchiseSubType (Тип франшизы)
                if (row[idxMap.franchiseSubtype]) {
                    xml += `    <FranchiseSubType>${escapeXml(row[idxMap.franchiseSubtype])}</FranchiseSubType>\n`;
                }
                // Title
                if (row[idxMap.title]) {
                    xml += `    <Title>${escapeXml(row[idxMap.title])}</Title>\n`;
                }
                // Description (с учётом возможного отсутствия символа "<")
                let desc = row[idxMap.description] || '';
                desc = desc.trim();
                if (desc.startsWith('p>')) {
                    desc = '<' + desc;
                }
                xml += '    <Description><![CDATA[' + desc + ']]></Description>\n';
                // Price
                if (row[idxMap.price]) {
                    xml += `    <Price>${escapeXml(row[idxMap.price])}</Price>\n`;
                }
                // ContactMethod
                if (row[idxMap.contactMethod]) {
                    xml += `    <ContactMethod>${escapeXml(row[idxMap.contactMethod])}</ContactMethod>\n`;
                }
                // FranchiseFee (Паушальный взнос) – значение ≥ 0
                let fee = row[idxMap.fee];
                if (fee === '' || fee == null) fee = '0';
                xml += `    <FranchiseFee>${escapeXml(String(fee))}</FranchiseFee>\n`;
                // FranchiseRoyalty (Роялти)
                let royalty = row[idxMap.royalty];
                if (royalty === '' || royalty == null) royalty = 'Нет';
                xml += `    <FranchiseRoyalty>${escapeXml(String(royalty))}</FranchiseRoyalty>\n`;
                // RoyaltyType
                if (row[idxMap.royaltyType]) {
                    xml += `    <RoyaltyType>${escapeXml(row[idxMap.royaltyType])}</RoyaltyType>\n`;
                }
                // FixedRoyalty
                if (row[idxMap.fixedRoyalty]) {
                    xml += `    <FixedRoyalty>${escapeXml(row[idxMap.fixedRoyalty])}</FixedRoyalty>\n`;
                }
                // PercentRoyalty
                if (row[idxMap.percentRoyalty]) {
                    xml += `    <PercentRoyalty>${escapeXml(row[idxMap.percentRoyalty])}</PercentRoyalty>\n`;
                }
                // Payback
                if (row[idxMap.payback]) {
                    xml += `    <Payback>${escapeXml(row[idxMap.payback])}</Payback>\n`;
                }
                // Support
                if (row[idxMap.support]) {
                    xml += `    <Support>${escapeXml(row[idxMap.support])}</Support>\n`;
                }
                // SupportType
                if (row[idxMap.supportType]) {
                    xml += `    <SupportType>${escapeXml(row[idxMap.supportType])}</SupportType>\n`;
                }
                // TargetAudience
                if (row[idxMap.audience]) {
                    xml += `    <TargetAudience>${escapeXml(row[idxMap.audience])}</TargetAudience>\n`;
                }
                // ContactEmail
                if (row[idxMap.email]) {
                    xml += `    <ContactEmail>${escapeXml(row[idxMap.email])}</ContactEmail>\n`;
                }
                // DateEnd
                if (row[idxMap.dateEnd]) {
                    xml += `    <DateEnd>${escapeXml(row[idxMap.dateEnd])}</DateEnd>\n`;
                }
                // VideoURL
                if (row[idxMap.video]) {
                    xml += `    <VideoURL>${escapeXml(row[idxMap.video])}</VideoURL>\n`;
                }
                // CompanyName
                if (row[idxMap.company]) {
                    xml += `    <CompanyName>${escapeXml(row[idxMap.company])}</CompanyName>\n`;
                }
                // Images
                const photosCell = row[idxMap.photos];
                if (photosCell) {
                    const urls = photosCell.split('|').map(u => u.trim()).filter(Boolean);
                    if (urls.length) {
                        xml += '    <Images>\n';
                        urls.forEach(url => {
                            xml += `      <Image url="${escapeXml(url)}" />\n`;
                        });
                        xml += '    </Images>\n';
                    }
                }
                xml += '  </Ad>\n';
            }
            xml += '</Ads>\n';
            // Создаём Blob и ссылку для скачивания
            const blob = new Blob([xml], { type: 'application/xml' });
            const downloadUrl = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = downloadUrl;
            link.download = 'avito_feed.xml';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(downloadUrl);
            statusBox.textContent = 'Готово! Файл XML сформирован и скачан.';
        } catch (err) {
            console.error(err);
            statusBox.textContent = 'Ошибка: ' + err.message;
        }
    };
    reader.onerror = (err) => {
        console.error(err);
        statusBox.style.display = 'block';
        statusBox.textContent = 'Не удалось прочитать файл.';
    };
    reader.readAsArrayBuffer(file);
});

// Функция для экранирования спецсимволов в XML (кроме содержимого CDATA)
function escapeXml(unsafe) {
    return String(unsafe)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}