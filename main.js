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

async function convertFile(e) {
	const fileInput = document.getElementById('fileInput');
	const buttonText = e.target.innerText;
	const file = fileInput.files[0];

	if (!file) {
		alert('Выберите excel file');
		return;
	}

	const loadStart = () => {
		e.target.innerHTML = '';
		e.target.setAttribute('aria-busy', true);
	};

	const loadFinish = () => {
		e.target.setAttribute('aria-busy', false);
		e.target.innerHTML = buttonText;
	};

	loadStart();

	const reader = new FileReader();

	reader.onloadend = async function () {
		try {
			const data = new Uint8Array(reader.result);
			const workbook = XLSX.read(data, { type: 'array' });
			const sheetName = workbook.SheetNames[0];
			const sheet = workbook.Sheets[sheetName];

			const jsonDataForm = XLSX.utils.sheet_to_json(sheet, { header: 1 });

			// Первая строка (индекс 0) содержит названия столбцов
			const headers = jsonDataForm[1];

			// Остальные строки содержат данные
			const dataRows = jsonDataForm.slice(2);
			// Преобразуем данные в массив объектов, используя названия столбцов в качестве ключей
			const translatedData = dataRows.map(row => {
				const obj = {};

				headers.forEach((header, index) => {
					// Применяем перевод только для заголовков, которые существуют в словаре translated
					const translatedWord = translated[header];
					obj[translatedWord || header] = row[index]; // Используем переведенный заголовок, если он есть
				});
				return obj;
			});

			const newTable = [...translatedData]
				.sort((a, b) => {
					const aN = a['Маркировка'] || 0;
					const bN = b['Маркировка'] || 0;
					if (!aN || !bN) return 0;
					return aN.localeCompare(bN);
				})
				.map(item => ({
					'Номер коробки': item['Номер коробки'],
					Маркировка: item['Маркировка'],
					'Номер накладной': item['Номер накладной'],
				}))
				.filter(item => Boolean(item['Маркировка']));

			const newSheet = XLSX.utils.json_to_sheet(newTable);
			const newWorkbook = XLSX.utils.book_new();
			XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
			const newBinaryData = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });

			const blob = new Blob([new Uint8Array(newBinaryData)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
			const blobURL = URL.createObjectURL(blob);

			const link = document.createElement('a');
			link.href = blobURL;
			link.download = 'Новая таблица заказов.xlsx';
			link.click();

			URL.revokeObjectURL(blobURL);
		} catch (error) {
			console.error('An error occurred:', error);
			alert('An error occurred. Please check the console for details.');
		} finally {
			loadFinish();
		}
	};

	reader.readAsArrayBuffer(file);
}
document.querySelector('#convertButton').addEventListener('click', convertFile);
