const xlsx = require('xlsx');
const workbook = xlsx.readFile('./test.xlsx');

const sheetNameList = workbook.SheetNames;
const promise = new Promise((resolve, reject) => {
    sheetNameList.forEach((y) => {
        const worksheet = workbook.Sheets[y];
        let headers = {};
        let data = [];
        for(z in worksheet) {
            if(z[0] === '!') continue;
            //parse out the column, row, and value
            var col = z.substring(0,1);
            var row = parseInt(z.substring(1));
            var value = worksheet[z].v;
    
            //store header names
            if(row == 1) {
                headers[col] = value;
                continue;
            }
    
            if(!data[row]) data[row]={};
            data[row][headers[col]] = value;
        }
        //drop those first two rows which are empty
        data.shift();
        data.shift();
        if(data){
            resolve(data);
        }else{
            reject("error");
        }
    }); 
});

promise.then(res => {
    console.log(res);
}).catch(err => {
    console.log(err);
});