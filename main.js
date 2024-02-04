const translated = {
	运输账单: 'счет за доставку',
	单号明细: 'Детали номера заказа',
	发货日期: 'Дата отправки',
	总包数: 'Общее количество упаковок',
	总重量: 'Общий вес',
	总立方: 'Общий объем',
	单价: 'Цена за единицу',
	包装费总计: 'Общая стоимость упаковки',
	总件数: 'Общее количество товаров',
	'应付货款（$）': 'Сумма к оплате ($)',
	客户名: 'Имя клиента',
	箱号: 'Номер коробки',
	收件日期: 'Дата получения',
	客户唛头: 'Маркировка',
	快递单号: 'Номер накладной',
	票号: 'Номер билета',
	包数: 'Количество упаковок',
	重量: 'Вес',
	立方: 'Объем',
	打包方式: 'Способ упаковки',
	包装费: 'Стоимость упаковки',
	小件数: 'Количество мелких товаров',
	目的地: 'Назначение',
};

function convertFile() {
	const fileInput = document.getElementById('fileInput');
	const jsonOutput = document.getElementById('jsonOutput');

	const file = fileInput.files[0];

	if (!file) {
		alert('Please select a file.');
		return;
	}

	const reader = new FileReader();

	reader.onload = function (e) {
		const data = new Uint8Array(e.target.result);
		const workbook = XLSX.read(data, { type: 'array' });
		const sheetName = workbook.SheetNames[0];
		const sheet = workbook.Sheets[sheetName];

		const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 'A' });

		// Применяем перевод только к ключам, которые существуют в translate
		const translatedData = jsonData.map(item => {
			const newTranslatedList = Object.entries(item).map(item => {
				const newTranslatedWord = translated[item[1]];
				if (newTranslatedWord) {
					return [item[0], newTranslatedWord];
				}
				return [...item];
			});

			return Object.fromEntries(newTranslatedList);
		});

		translatedData.sort((a, b) => {
			const aN = a.N || 0; // Если ключ "N" отсутствует, предполагаем значение 0
			const bN = b.N || 0;
			if (!aN || !bN) return 0;
			return aN.localeCompare(bN); // Сортировка как строки (если значения числовые, используйте aN - bN)
		});
		console.log(translatedData);
		jsonOutput.innerHTML = JSON.stringify(translatedData);
		// Создаем новый лист XLSX из переведенных данных
		const newSheet = XLSX.utils.json_to_sheet(translatedData);
		// Создаем новую книгу и добавляем лист
		const newWorkbook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Новая таблица');
		// Преобразуем книгу в бинарные данные XLSX
		const newBinaryData = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });

		// Создаем Blob из бинарных данных и формируем URL
		// const blob = new Blob([new Uint8Array(newBinaryData)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
		// const blobURL = URL.createObjectURL(blob);

		// // Устанавливаем ссылку на скачивание
		// const link = document.createElement('a');
		// link.href = blobURL;
		// link.download = 'translated_data.xlsx';

		// // Эмулируем клик по ссылке
		// link.click();

		// // Освобождаем URL
		// URL.revokeObjectURL(blobURL);
	};

	reader.readAsArrayBuffer(file);
}
