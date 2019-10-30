import helper from './helper';
import { expr2expr } from './alphabet';

class Rows {
  constructor({ len, height }) {
    this._ = {};
    this.len = len;
    // default row height
    this.height = height;
  }

  getHeight(ri) {
    const row = this.get(ri);
    if (row && row.height) {
      return row.height;
    }
    return this.height;
  }

  setHeight(ri, v) {
    const row = this.getOrNew(ri);
    row.height = v;
  }

  setStyle(ri, style) {
    const row = this.getOrNew(ri);
    row.style = style;
  }

  sumHeight(min, max, exceptSet) {
    return helper.rangeSum(min, max, (i) => {
      if (exceptSet && exceptSet.has(i)) return 0;
      return this.getHeight(i);
    });
  }

  totalHeight() {
    return this.sumHeight(0, this.len);
  }

  get(ri) {
    return this._[ri];
  }

  getOrNew(ri) {
    this._[ri] = this._[ri] || { cells: {} };
    return this._[ri];
  }

  getCell(ri, ci) {
    const row = this.get(ri);
    if (row !== undefined && row.cells !== undefined && row.cells[ci] !== undefined) {
      return row.cells[ci];
    }
    return null;
  }

  getCellMerge(ri, ci) {
    const cell = this.getCell(ri, ci);
    if (cell && cell.merge) return cell.merge;
    return [0, 0];
  }

  getCellOrNew(ri, ci) {
    const row = this.getOrNew(ri);
    row.cells[ci] = row.cells[ci] || {};
    return row.cells[ci];
  }

  // what: all | text | format
  setCell(ri, ci, cell, what = 'all') {
    const row = this.getOrNew(ri);
    if (what === 'all') {
      row.cells[ci] = cell;
    } else if (what === 'text') {
      row.cells[ci] = row.cells[ci] || {};
      row.cells[ci].text = cell.text;
    } else if (what === 'format') {
      row.cells[ci] = row.cells[ci] || {};
      row.cells[ci].style = cell.style;
      if (cell.merge) row.cells[ci].merge = cell.merge;
    }
  }

  setCellText(ri, ci, text) {
    const cell = this.getCellOrNew(ri, ci);
    cell.text = text;
  }

  // what: all | format | text
  copyPaste(srcCellRange, dstCellRange, what, autofill = false, cb = () => {}) {
    const {
      sri, sci, eri, eci,
    } = srcCellRange;
    const dsri = dstCellRange.sri;
    const dsci = dstCellRange.sci;
    const deri = dstCellRange.eri;
    const deci = dstCellRange.eci;
    const [rn, cn] = srcCellRange.size();
    const [drn, dcn] = dstCellRange.size();
    // console.log(srcIndexes, dstIndexes);
    let isAdd = true;
    let dn = 0;
    if (deri < sri || deci < sci) {
      isAdd = false;
      if (deri < sri) dn = drn;
      else dn = dcn;
    }
    // console.log('drn:', drn, ', dcn:', dcn, dn, isAdd);
    for (let i = sri; i <= eri; i += 1) {
      if (this._[i]) {
        for (let j = sci; j <= eci; j += 1) {
          if (this._[i].cells && this._[i].cells[j]) {
            for (let ii = dsri; ii <= deri; ii += rn) {
              for (let jj = dsci; jj <= deci; jj += cn) {
                const nri = ii + (i - sri);
                const nci = jj + (j - sci);
                const ncell = helper.cloneDeep(this._[i].cells[j]);
                // ncell.text
                if (autofill && ncell && ncell.text && ncell.text.length > 0) {
                  const { text } = ncell;
                  let n = (jj - dsci) + (ii - dsri) + 2;
                  if (!isAdd) {
                    n -= dn + 1;
                  }
                  if (text[0] === '=') {
                    ncell.text = text.replace(/\w{1,3}\d/g, (word) => {
                      let [xn, yn] = [0, 0];
                      if (sri === dsri) {
                        xn = n - 1;
                        // if (isAdd) xn -= 1;
                      } else {
                        yn = n - 1;
                      }
                      // console.log('xn:', xn, ', yn:', yn, word, expr2expr(word, xn, yn));
                      return expr2expr(word, xn, yn);
                    });
                  } else {
                    const result = /[\\.\d]+$/.exec(text);
                    // console.log('result:', result);
                    if (result !== null) {
                      const index = Number(result[0]) + n - 1;
                      ncell.text = text.substring(0, result.index) + index;
                    }
                  }
                }
                // console.log('ncell:', nri, nci, ncell);
                this.setCell(nri, nci, ncell, what);
                cb(nri, nci, ncell);
              }
            }
          }
        }
      }
    }
  }

  cutPaste(srcCellRange, dstCellRange) {
    const ncellmm = {};
    this.each((ri) => {
      this.eachCells(ri, (ci) => {
        let nri = parseInt(ri, 10);
        let nci = parseInt(ci, 10);
        if (srcCellRange.includes(ri, ci)) {
          nri = dstCellRange.sri + (nri - srcCellRange.sri);
          nci = dstCellRange.sci + (nci - srcCellRange.sci);
        }
        ncellmm[nri] = ncellmm[nri] || { cells: {} };
        ncellmm[nri].cells[nci] = this._[ri].cells[ci];
      });
    });
    this._ = ncellmm;
  }

  insert(sri, n = 1) {
    const ndata = {};
    this.each((ri, row) => {
      let nri = parseInt(ri, 10);
      this.eachCells(ri, (ci, cell) => {
        if(cell['text'] && cell['text'].startsWith("=")){  // checking if it is a formula
          cell['text'] = cell['text'].substr(1)            // removing "=" from the formula string
          let cellText  = cell['text']
          let numberPattern = /[A-Z]\d+/g;                 // It will find all numbers preceded by a charcter Eg: 26 from B26
          let numbersArr = cellText.match(numberPattern)
          if(numbersArr){
            numbersArr = numbersArr.reverse()
            numbersArr.map(value=>{
                let oldIndex = parseInt(value.substr(1))        // we get the index that needs to be updated Eg: 26 from B26
                if(oldIndex >= (sri+1)){                         // If index is greater than the deleted row index then we updated the refernece
                  let updatedIndexStr = value[0] + (oldIndex + n);
                  let updatedCell = cell['text'].replace(new RegExp("\\b"+value+"\\b"),updatedIndexStr)
                  cell['text'] = updatedCell
                }
              })
          }
          cell['text'] = "="+cell['text']
        }
      })
      if (nri >= sri) {
        nri += n;
      }
      ndata[nri] = row;
    });
    this._ = ndata;
    this.len += n;
  }

  delete(sri, eri) {
    const n = eri - sri + 1;
    const ndata = {};
    this.each((ri, row) => {
      const nri = parseInt(ri, 10);
      this.eachCells(ri, (ci, cell) => {
        if(cell['text'] && cell['text'].startsWith("=")){  // checking if it is a formula
          cell['text'] = cell['text'].substr(1)            // removing "=" from the formula string
          let cellText  = cell['text']
          //console.log("Before updating :: ",cell['text'])
          let rangePattern = /[A-Z]\d+[:][A-Z]\d+/g;       // It will find all range patterns Eg: B23:B26
          let numberPattern = /[A-Z]\d+/g;                 // It will find all numbers preceded by a charcter Eg: 26 from B26
          let rangeValuesArr = cellText.match(rangePattern)
          if(rangeValuesArr){
            rangeValuesArr.map(value=>{
              cellText = cellText.replace(value,'')         // once we get a range pattern we remove it from formula string so it won't be evaluated agin in next loop
              let rangeNumbers = value.match(numberPattern)
              let startRange = rangeNumbers[0]              // rangeNumbers is a array of two numbers of range Eg: [23,26] from B23:B26
              let endRange = rangeNumbers[1]
              let startRangeIndex = parseInt(startRange.substr(1))
              let endRangeIndex = parseInt(endRange.substr(1))
              if(endRangeIndex == (eri+1) && startRangeIndex == (sri+1)){         // If we delete whole range then we add #REF! error in formula
                let updatedCell = cell['text'].replace(new RegExp("\\b"+startRange+"\\b"),'"#REF!"')
                cell['text'] = updatedCell
                updatedCell = cell['text'].replace(new RegExp("\\b"+endRange+"\\b"),'"#REF!"')
                cell['text'] = updatedCell
              }else if(endRangeIndex > (eri+1) || (endRangeIndex <= (eri+1) && endRangeIndex >= (sri+1))){ // If we delete some part of range then we update the range end index
                let updatedIndexStr = endRange[0] + (endRangeIndex - n);
                let updatedCell = cell['text'].replace(new RegExp("\\b"+endRange+"\\b"),updatedIndexStr)
                cell['text'] = updatedCell
              }
              if(startRangeIndex > (eri+1) ||(startRangeIndex <= (eri+1) && startRangeIndex >= (sri+1))){
                let updatedIndexStr = startRange[0] + (startRangeIndex - n);
                let updatedCell = cell['text'].replace(new RegExp("\\b"+startRange+"\\b"),updatedIndexStr)
                cell['text'] = updatedCell
              }
            })
          }
          
          let numbersArr = cellText.match(numberPattern)
          if(numbersArr){
            numbersArr.map(value=>{
                let oldIndex = parseInt(value.substr(1))        // we get the index that needs to be updated Eg: 26 from B26
                if(oldIndex > (eri+1)){                         // If index is greater than the deleted row index then we updated the refernece
                  let updatedIndexStr = value[0] + (oldIndex - n);
                  let updatedCell = cell['text'].replace(new RegExp("\\b"+value+"\\b"),updatedIndexStr)
                  cell['text'] = updatedCell
                }else if(oldIndex <= (eri+1) && oldIndex >= (sri+1)){
                  let updatedCell = cell['text'].replace(new RegExp("\\b"+value+"\\b"),'"#REF!"')
                  cell['text'] = updatedCell
                }
              })
          }
          cell['text'] = "="+cell['text']
          //console.log("After updating ::", cell['text'])
        }
      })

      if (nri < sri) {
        ndata[nri] = row;
      } else if (nri > eri) {
        ndata[nri - n] = row;
      }
    });
    this._ = ndata;

    this.len -= n;
  }

  insertColumn(sci, n = 1) {
    this.each((ri, row) => {
      const rndata = {};
      this.eachCells(ri, (ci, cell) => {
        let nci = parseInt(ci, 10);

        if(cell['text'] && cell['text'].startsWith("=")){  // checking if it is a formula
          cell['text'] = cell['text'].substr(1)            // removing "=" from the formula string
          let cellText  = cell['text']
          let numberPattern = /[A-Z]\d+/g;                 // It will find all numbers preceded by a charcter Eg: 26 from B26
          let numbersArr = cellText.match(numberPattern)
          if(numbersArr){
            numbersArr.map(value=>{
                let oldColIndex = (value[0].charCodeAt(0) - 65) +1;
                if(oldColIndex >= (sci+1)){                         // If index is greater than the deleted row index then we updated the refernece
                  let updatedIndexStr = String.fromCharCode((oldColIndex-1+n+65)) + value.substr(1);
                  let updatedCell = cell['text'].replace(new RegExp("\\b"+value+"\\b"),updatedIndexStr)
                  cell['text'] = updatedCell
                }
            })
          }
          cell['text'] = "="+cell['text']
        }


        if (nci >= sci) {
          nci += n;
        }
        rndata[nci] = cell;
      });
      row.cells = rndata;
    });
  }

  deleteColumn(sci, eci) {
    const n = eci - sci + 1;
    this.each((ri, row) => {
      const rndata = {};
      this.eachCells(ri, (ci, cell) => {
        const nci = parseInt(ci, 10);
        if(cell['text'] && cell['text'].startsWith("=")){  // checking if it is a formula
          cell['text'] = cell['text'].substr(1)            // removing "=" from the formula string
          let cellText  = cell['text']
          let numberPattern = /[A-Z]\d+/g;                 // It will find all numbers preceded by a charcter Eg: 26 from B26
          let numbersArr = cellText.match(numberPattern)
          if(numbersArr){
            numbersArr.map(value=>{
                let oldColIndex = (value[0].charCodeAt(0) - 65) +1;
                if(oldColIndex > (eci+1)){                         // If index is greater than the deleted row index then we updated the refernece
                  let updatedIndexStr = String.fromCharCode((oldColIndex-1-n+65)) + value.substr(1);
                  let updatedCell = cell['text'].replace(new RegExp("\\b"+value+"\\b"),updatedIndexStr)
                  cell['text'] = updatedCell
                }else if(oldColIndex <= (eci+1) && oldColIndex >= (sci+1)){
                  let updatedCell = cell['text'].replace(new RegExp("\\b"+value+"\\b"),'"#REF!"')
                  cell['text'] = updatedCell
                }
            })
          }
          cell['text'] = "="+cell['text']
        }
        if (nci < sci) {
          rndata[nci] = cell;
        } else if (nci > eci) {
          rndata[nci - n] = cell;
        }
      });
      row.cells = rndata;
    });
  }

  // what: all | text | format | merge
  deleteCells(cellRange, what = 'all') {
    cellRange.each((i, j) => {
      this.deleteCell(i, j, what);
    });
  }

  // what: all | text | format | merge
  deleteCell(ri, ci, what = 'all') {
    const row = this.get(ri);
    if (row !== null) {
      const cell = this.getCell(ri, ci);
      if (cell !== null) {
        if (what === 'all') {
          delete row.cells[ci];
        } else if (what === 'text') {
          if (cell.text) delete cell.text;
          if (cell.value) delete cell.value;
        } else if (what === 'format') {
          if (cell.style !== undefined) delete cell.style;
          if (cell.merge) delete cell.merge;
        } else if (what === 'merge') {
          if (cell.merge) delete cell.merge;
        }
      }
    }
  }

  each(cb) {
    Object.entries(this._).forEach(([ri, row]) => {
      cb(ri, row);
    });
  }

  eachCells(ri, cb) {
    if (this._[ri] && this._[ri].cells) {
      Object.entries(this._[ri].cells).forEach(([ci, cell]) => {
        cb(ci, cell);
      });
    }
  }

  setData(d) {
    console.log(d)
    if (d.len) {
      this.len = d.len;
      delete d.len;
    }
    this._ = d;
  }

  getData() {
    const { len } = this;
    return Object.assign({ len }, this._);
  }
}

export default {};
export {
  Rows,
};
