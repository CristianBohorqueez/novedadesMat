// let firstDate = new Date(2022, 12 - 1, 28)
// let finalDate = new Date(2023, 1 - 1, 12);
// let daysOfYear = [];
// for (finalDate; firstDate <= finalDate; firstDate.setDate(firstDate.getDate() + 1)) {
//     daysOfYear.push(new Date(firstDate).toLocaleDateString("es-MX"));
// }

class Excel {
  constructor(_data) {
    this._data = _data;
  }
  header() {
    return this._data[0];
  }
  rows() {
    return new RowCollection(this._data.slice(1, this._data.length));
  }
}

class RowCollection {
  constructor(rows) {
    this.rows = rows;
  }
  getFirst() {
    return new Row(this.rows[1]);
  }
  get(index) {
    return new Row(this.rows[index]);
  }
  count() {
    return this.rows.length;
  }
}

class Row {
  constructor(_row) {
    this._row = _row;
  }
  radicado() {
    let dateRadicado = new Date (this._row[0].toLocaleDateString("en-US"))
    dateRadicado =  (dateRadicado.getDate() + 1) + '/' + (dateRadicado.getMonth() + 1)  + '/' + dateRadicado.getFullYear();
    return dateRadicado;
  }
  identificacion() {
    return this._row[1];
  }
  nombre() {  
    return this._row[2];
  }
  motivo() {  
    return this._row[3];
  }
  fechaInicio() {  
    let dateRadicado = new Date (this._row[4].toLocaleDateString("en-US"))
    dateRadicado =  (dateRadicado.getDate()) + '/' + (dateRadicado.getMonth() + 1)  + '/' + dateRadicado.getFullYear();
    return dateRadicado;
  }
  fechaInicioOld() {
    return this._row[4];
  }
  fechaFin() {  
    let dateRadicado = new Date (this._row[5].toLocaleDateString("en-US"))
    dateRadicado =  (dateRadicado.getDate()) + '/' + (dateRadicado.getMonth() + 1)  + '/' + dateRadicado.getFullYear();
    return dateRadicado;
  }
  fechaFinOld() {
    return this._row[5];
  }
  dias() {  
    return this._row[6];
  }
  ceco() {  
    return this._row[7];
  }
  numeroCeco() {  
    return this._row[8];
  }
}

class ExcelPrinter {
	static print(tableId, excel) {
		let table = document.getElementById(tableId);

		excel.header().forEach((title) => {
			table.querySelector('thead>tr').innerHTML += `<th>${title}</th>`;
		});

		for (let i = 0; i < excel.rows().count(); i++) {
			let row = excel.rows().get(i);
			let tr = document.createElement('tr');
			tr.innerHTML = `<td>${row.radicado()}</td>
                      <td>${row.identificacion()}</td>
                      <td>${row.nombre()}</td>
                      <td>${row.motivo()}</td>
                      <td>${row.fechaInicio()}</td>
                      <td>${row.fechaFin()}</td>
                      <td>${row.dias()}</td>
                      <td>${row.ceco()}</td>
                      <td>${row.numeroCeco()}</td>`;
			table.querySelector('tbody').appendChild(tr);
		}
	}
}

const printRangeData = (excel) => {
  let table = document.getElementById('table_content');
  table.querySelector('thead>tr').innerHTML = '';
  table.querySelector('tbody').innerHTML = '';
  for (let i = 0; i < excel.rows().count(); i++) {
    let row = excel.rows().get(i);
    let firstDate = row.fechaInicioOld()
    let finalDate = row.fechaFinOld()
    let rangeData = [];
    firstDate.setDate(firstDate.getDate() + 1);
    finalDate.setDate(finalDate.getDate() + 1);
    for (finalDate; firstDate <= finalDate; firstDate.setDate(firstDate.getDate() + 1)) {
        rangeData.push(new Date(firstDate).toLocaleDateString("es-MX"));
        let tr = document.createElement('tr');
        tr.innerHTML = `<td>${row.radicado()}</td>
                        <td>${row.identificacion()}</td>
                        <td>${row.nombre()}</td>
                        <td>${row.motivo()}</td>
                        <td>${row.fechaInicio()}</td>
                        <td>${row.fechaFin()}</td>
                        <td>${row.dias()}</td>
                        <td>${row.ceco()}</td>
                        <td>${row.numeroCeco()}</td>`;
        table.querySelector('tbody').appendChild(tr);
    }
  }
  console.log(excel, 'rangeData');
}

const htmlTableToExcel = async (type) => {
  const fileName = 'novedades_Mat.' + type;
  const table = document.getElementById('table_content');
  const wb = XLSX.utils.table_to_book(table);
  XLSX.writeFile(wb, fileName);
 }

function exportCSVExcel() {
  var tableElement = document.getElementById('table_content');
  var sourceData = 'data:text/csv;charset=utf-8,';
  var i = 0;
  while ((row = tableElement.rows[i])) {
    sourceData +=
      [
        row.cells[0],
        row.cells[1],
        row.cells[2],
        row.cells[3],
      ].join(',') + '\r\n';
    i++;
  }
  window.location.href = encodeURI(sourceData);
}

const input = document.getElementById('file-selector')
  input.addEventListener('change',async function() {
    const content = await readXlsxFile(input.files[0])
    const excel = new Excel(content)
    printRangeData(excel)
    // ExcelPrinter.print('table_content', excel);
  })
