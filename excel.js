let exFile;
document.getElementById('input').addEventListener("change", (event) => {
    exFile = event.target.files[0];
})
document.getElementById('button').addEventListener("click", () => {
    
    if(exFile){
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(exFile);
        fileReader.onload = (event)=>{
         let data = event.target.result;
         let workbook = XLSX.read(data,{type:"binary"});
         workbook.SheetNames.forEach(sheet => {
              let exSheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);
              //let text="";
              let fail=0;
              let success=0;
              let resi=0;
              var barcodes = document.getElementById("jsondata");
              for (var j=0; j<=exSheet.length-1; j++){
                resi+=1
                let rowSheet = exSheet[j]
                if(Object.entries(rowSheet)[0][0]!="No. Resi"){
                    fail+=1;
                    continue;
                }
                else{
                    success+=1;
                    let awb = Object.entries(rowSheet)[0][1]
                    let recepient = Object.entries(rowSheet)[9][1]
                    var newBarcode = document.createElementNS("http://www.w3.org/2000/svg", "svg");
                    JsBarcode(newBarcode, awb,{
                        text: awb+`\n`+recepient.slice(0,10),
                        width:2.5,
                        margin: 15,
                        fontSize: 18
                    });
                    barcodes.appendChild(newBarcode);
                }
              }
              if(success>0){
                    document.getElementById("fail").innerHTML = "<br>Jumlah resi: "+resi+"<br>Jumlah resi gagal cetak: "+fail+`<br><button class="btn btn-success" onclick="window.print()">Print Resi</button>`;
              }
              else{
                document.getElementById("fail").innerHTML = "<br>Maaf, resi gagal dicetak<br>Jumlah resi: "+resi+"<br>Jumlah resi gagal cetak: "+fail;
              }
              
         });
        }
    }
});
function printDiv(printable) {
    var printContents = document.getElementById(printable).innerHTML;
    var originalContents = document.body.innerHTML;

    document.body.innerHTML = printContents;

    window.print();

    document.body.innerHTML = originalContents;
}