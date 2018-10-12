
var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer

function getJsDateFromExcel(excelDate) {

  // JavaScript dates can be constructed by passing milliseconds
  // since the Unix epoch (January 1, 1970) example: new Date(12312512312);

  // 1. Subtract number of days between Jan 1, 1900 and Jan 1, 1970, plus 1 (Google "excel leap year bug")             
  // 2. Convert to milliseconds.
  return new moment((excelDate - (25567 + 2))*86400*1000);
}

class Movement extends Object{
  constructor(cells, baseCellName){
    super();
    let recipientCellRow = parseInt( baseCellName.slice(2) ) + 2;
    let recipientCellName = 'K' + recipientCellRow;
    
    let okpoCellName = 'AK' + recipientCellRow;

    let infoCellRow = parseInt( baseCellName.slice(2) ) + 3;
    let infoCellName = 'J' + infoCellRow;

    let dateCellRow = parseInt( baseCellName.slice(2) ) + 1;
    let dateCellName = 'AM' + dateCellRow;
    let date = getJsDateFromExcel( cells[dateCellName].v )


    this.data = {
      addr: baseCellName,
      sum: parseFloat(cells[baseCellName].v.replace(/,/, '.')),
      agent: cells[recipientCellName] ? cells[recipientCellName].v : 'None',
      agentEdrpou: cells[okpoCellName] ? cells[okpoCellName].v.slice(5) : 'None',
      info: cells[infoCellName] ? cells[infoCellName].v : 'None',
      date: date.format("DD.MM.YYYY")
    }
  }
}


class  Movements extends Object{
  constructor(oschadStruct){
    super();
    this.creditColumn = 'AA'
    this.debitColumn = 'AJ'
    this.allCredit = this.getFromOschad(oschadStruct, this.creditColumn);
    this.allDebit = this.getFromOschad(oschadStruct, this.debitColumn);
  }

  getFromOschad(w, valueColumn) {
    const cells = w.Sheets[ w.SheetNames[0] ];
    const res = new Array;

    for( let cellName in cells ){

      if (cellName.slice(0,2) == valueColumn) {
        let cellVal = parseFloat(cells[cellName].v.replace(/,/, '.'))

        if (!isNaN(cellVal)) {
          let movement = new Movement(cells, cellName) 
          res.push(movement)
        }
      }
    }
    return res;
  }

  drawTo(jqET){
    jqET.html('')
    jqET.append('<thead></thead>')
      .find('thead')
      .append('<th>date</th>')
      .append('<th>income UAH</th>')
      .append('<th>income USD</th>')
      .append('<th>outcome UAH</th>')
      .append('<th>outcome USD</th>')
      .append('<th>agent</th>')
      .append('<th>detail</th>')

    const jqBody = jqET.append('<tbody></tbody>').find('tbody')

    this.allCredit.forEach( (ac) =>{
      jqBody.append('<tr></tr>').find('tr').last()
        .append('<td>'+ac.data.date+'</td>')
        .append('<td class="money"> 0,00 </td>')
        .append('<td class="money"> 0,00 </td>')
        .append('<td class="money">'+ac.data.sum.toFixed(2).replace(/\./, ',')+ '</td>')
        .append('<td class="money"> 0,00 </td>')
        .append('<td>'+ac.data.agent+'</td>')
        .append('<td>'+ac.data.info+'</td>')

///
      // addr: baseCellName,
      // sum: parseFloat(cells[baseCellName].v.replace(/,/, '.')),
      // agent: cells[recipientCellName] ? cells[recipientCellName].v : 'None',
      // agentEdrpou: cells[okpoCellName] ? cells[okpoCellName].v.slice(5) : 'None',
      // info: cells[infoCellName] ? cells[infoCellName].v : 'None',
      // date: date.format("DD.MM.YYYY")
///      



    })





  }

}


var reader = new FileReader();
reader.onload = function(e) {
  var data = e.target.result;
  if(!rABS) data = new Uint8Array(data);
  var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});

  /* DO SOMETHING WITH workbook HERE */    
  let movements = new Movements( workbook )
  console.log( movements )
  const jqElementTable = $('table#result')

  movements.drawTo(jqElementTable)
  /* DO SOMETHING WITH workbook HERE */

};


function handleDrop(e) {
	console.log('handleDrop!')
  e.stopPropagation(); e.preventDefault();
  var files = e.dataTransfer.files, f = files[0];
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}


function handleDragover(e) {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}


function handleFile(e) {
  console.log('handleFile!')  
  var files = e.target.files, f = files[0];
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}



const drop = document.getElementById('drop-area')
drop.addEventListener('dragenter', handleDragover, false);
drop.addEventListener('dragover', handleDragover, false);
drop.addEventListener('drop', handleDrop, false);

 
const xlf = document.getElementById('xlf');
xlf.addEventListener('change', handleFile, false);
