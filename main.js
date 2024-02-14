'use strict';
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
	try {
		const fileInput = document.getElementById('fileInput');
		const buttonText = e.target.innerText;
		e.target.setAttribute('aria-invalid', false);

		const loadStart = () => {
			e.target.innerHTML = '';
			e.target.setAttribute('aria-busy', true);
		};
		loadStart();
		const loadFinish = () => {
			e.target.setAttribute('aria-busy', false);
			e.target.innerHTML = buttonText;
		};

		const file = fileInput.files[0];
		if (!file) {
			alert('Выберете файл!');
		}
		const reader = new FileReader();

		reader.onloadend = () => {
			loadFinish();
		};

		reader.onload = function (event) {
			const data = event.target.result;
			const wb = XLSX.read(new Uint8Array(data), { type: 'array', bookVBA: true });
			const sheetNames = wb.SheetNames;
			const newWorkbook = XLSX.utils.book_new();

			sheetNames.forEach(sheetName => {
				const sheet = wb.Sheets[sheetName];
				const jsonDataForm = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
				const headers = jsonDataForm[1];
				const dataRows = jsonDataForm.slice(2);

				const translatedData = dataRows.map(row => {
					const obj = {};

					headers.forEach((header, index) => {
						const translatedWord = translated[header];
						obj[translatedWord || header] = row[index];
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
				XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
			});

			const newBinaryData = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });

			const blob = new Blob([new Uint8Array(newBinaryData)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
			const blobURL = URL.createObjectURL(blob);

			const link = document.createElement('a');
			link.href = blobURL;
			link.download = 'Новая таблица заказов.xlsx';
			document.body.appendChild(link);
			link.click();
		};

		reader.readAsArrayBuffer(file);
	} catch (error) {
		console.error(error);
		e.target.setAttribute('aria-invalid', true);
		const p = document.createElement('p');
		p.append(error);
		p.style.color = 'red';
		document.body.querySelector('.container').append(p);
	}
}

document.querySelector('#convertButton').addEventListener('click', convertFile);

document.querySelector('#convertButton').addEventListener('click', convertFile);
