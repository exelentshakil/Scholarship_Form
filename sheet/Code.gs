function JOINRANGES(range1, index1, range2, index2) {
 const result = []
 for(let row1 of range1) {
   for (let row2 of range2) {
    //google sheet based number, 1 based indexing
     if (row1[index1-1] == row2[index2-1]) {
       //console.log(row1,row1[index1-1],row2,row2[index2-1],(row1[index1-1] == row2[index2-1]))
       const r = [...row1, ...row2]
       // Remove the keys themselves
       r.splice(row1.length+(index2-1), 1)
       r.splice((index1-1), 1)
       result.push(r)
     }
   }
 }
 return result
}
function LEFTJOINRANGES(range1, index1, range2, index2,removejoin) {
 const result = []
 for(let row1 of range1) {
   isRowUsed=false
   lastrow2= [];
   for (let row2 of range2) {
    //google sheet based number, 1 based indexing
     if (row1[index1-1] == row2[index2-1]) {
       isRowUsed=true
       //console.log(row1,row1[index1-1],row2,row2[index2-1],(row1[index1-1] == row2[index2-1]))
       const r = [...row1, ...row2]
       // Remove the keys themselves
       
       if (removejoin[0][1]==1){
       r.splice(row1.length+(index2-1), 1)
       }
       if (removejoin[0][0]==1){
       r.splice((index1-1), 1)
       }
       result.push(r)
       lastrow2= row2;
     }
   }
   if(!isRowUsed){
     const r = [...row1, ...lastrow2.map(ee=>''),]
     
     if (removejoin[0][1]==1){
     r.splice(row1.length+(index2-1), 1)}
     if (removejoin[0][0]==1){
     r.splice((index1-1), 1)
     }
     result.push(r)
   }
 }
 return result
}

function JOINRANGESContained(range1, index1, range2, index2) {
 const result = []
 for(let row1 of range1) {
   for (let row2 of range2) {
    //google sheet based number, 1 based indexing
     if (row2[index2-1].indexOf( row1[index1-1])>-1 && row2[index2-1] && row1[index1-1] ) {
       //console.log(row1,row1[index1-1],row2,row2[index2-1],(row2[index2-1].indexOf( row1[index1-1])>-1))
       const r = [...row1, ...row2]
       // Remove the keys themselves
       r.splice(row1.length+(index2-1), 1)
       //r.splice((index1-1), 1)
       result.push(r)
     }
   }
 }
 return result
}


function JOINMonthlyFilteredRANGESContained(range1, index1, range2, index2,filteredIndex,range3, index3row, index3col) {
 // var ss = SpreadsheetApp.getActiveSpreadsheet();
//var sheet = ss.getSheets()[0];
 const result = []
 const months = ['01','02','03','04','05','06','07','08','09','10','11','12']
 var monthNames = [ "January", "February", "March", "April", "May", "June",
"July", "August", "September", "October", "November", "December" ];
 const quarters = ['Q1','Q1','Q1','Q2','Q2','Q2','Q3','Q3','Q3','Q4','Q4','Q4']
 for(let row1 of range1) {
   for (let row2 of range2) {
    //google sheet based number, 1 based indexing
     if (row2[index2-1].indexOf( row1[index1-1])>-1 && row2[index2-1] && row1[index1-1] ) {
       
       const r = [...row1, ...row2]
       // Remove the keys themselves
       r.splice(row1.length+(index2-1), 1)
       //r.splice((index1-1), 1)
       for (let [mi, month] of months.entries()) {  
          rm = JSON.parse(JSON.stringify(r))
          rm.push(month)
          rm.push(rm[index3row-1]+'-'+row1[index1-1]+'-'+month)
          rm.push(quarters[mi])
          if (rm[filteredIndex-1].indexOf(monthNames[mi])>-1 || monthNames.every(a=>rm[filteredIndex-1].indexOf(a)==-1))
          {
            //var cell = sheet.getRange("P"+(result.length+2).toString());
            //cell.setFormulaR1C1('=INDEX(XLOOKUP(G2, Submitted!A1:A,Submitted!A1:BZ),1,XMATCH(D2, Submitted!A1:BZ1))');
            //rm[index3row]
            subrow = range3.filter(r =>r[0]==rm[index3row-1])[0]
            subcol = range3[0].indexOf(rm[index3col-1])

            
            rm.push(monthNames[mi])
            var toPush=  true;
            if (subrow){
              //rm.push(subrow[subcol])
              
              toPush= subrow[subcol].toLowerCase() !='no'
              if (toPush){
                //rm.push((!monthNames.every(a=>subrow[subcol].toLowerCase().indexOf(a.toLowerCase())==-1)).toString())
                //rm.push(((subrow[subcol].toLowerCase().indexOf(monthNames[mi].toLowerCase())==-1)).toString())
                if(!monthNames.every(a=>subrow[subcol].toLowerCase().indexOf(a.toLowerCase())==-1)&&(subrow[subcol].toLowerCase().indexOf(monthNames[mi].toLowerCase())==-1))
                {
                  toPush=  false;
                }
                
                if(toPush && !quarters.every(a=>subrow[subcol].toLowerCase().indexOf(a.toLowerCase())==-1)&&(subrow[subcol].toLowerCase().indexOf(quarters[mi].toLowerCase())==-1))
                {
                  toPush=  false;
                }
                      
              }
            }else{
              console.log(rm[index3col-1],subcol,subrow,index3col)
            }
            if (toPush)
            {
              result.push(rm);
            }  
            
          }
       }
     }
   }
 }
 return result
}
function onEdit(e){
  if (!e){
    console.log("Do not run manually")
    throw "Do not run manually";
  } 
  
  syncNoID(e);
}

function syncNoID(e){
  const src = e.source.getActiveSheet();
  const r = e.range;
   var monthNames = [ "Jan", "Feb", "March", "April", "May", "June",
"July", "August", "September", "October", "November", "December" ];
  if (monthNames.indexOf(src.getName()) == -1 || r.rowStart == 1 || r.getColumn()<=8) return;
  
  let id = src.getRange(r.rowStart,4,1,10).getValues()[0][0];//src.getRange(r.rowStart,4,1,1).getValue().toString();
  //console.log('moni',src.getName(),id)
  //console.log("D"+r.rowStart,id);
  let newvalue =r.getValue().toString()
  
  r.clear();

  
  const db = SpreadsheetApp.getActive().getSheetByName("markdb");
  let dbData = db.getDataRange().getValues().map(e=>e[0]);
  console.log(dbData);
  for (let i in dbData){
    console.log(i,dbData);
    if (dbData[i].toString() == id){
      
      console.log('in2',i+1,dbData[i],id,r.getColumn(),r.getColumn()-7)
      celled= db.getRange(i+1,r.getColumn()-7,1,1)
      console.log('in3',newvalue,celled)
      celled.setValue(newvalue);
      return;
    }
  }
  var lastRowInd = db.getLastRow()+1;
  console.log(lastRowInd)
  db.getRange(lastRowInd,r.getColumn()-7,1,1).setValue(newvalue);
  db.getRange(lastRowInd,1,1,1).setValue(id);

}
