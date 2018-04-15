function reportData(name, http) {
    var self = this;
    self.name = name;
    self.link = http;
}


function reportViewModel(){

    this.reportList = ko.observableArray([
        new reportData("k2s", "http://47.88.132.185/api/v2/k2s_report" ),
        new reportData("Pam", "http://47.88.132.185/api/v2/k2s_report" ),
    ]);


    var selectedReport = ko.observable('');
    this.selectedReport = selectedReport();


function doit(type, fn, dl) {
	var elt = document.getElementById('data-table');
    var wb = XLSX.utils.table_to_book(elt, {sheet:"asdas"});
    console.log(elt);
    wb['!cols']=[{'wpx':500},{'wpx':500},{'wpx':500},{'wpx':500},{wpx:500},{wpx:500},{wpx:500},{wpx:500}];
    console.log(wb);

   
	return dl ?
		XLSX.write(wb, {bookType:type, bookSST:true, type: 'base64'}) :
		XLSX.writeFile(wb, fn || ('test.' + (type || 'xlsx'))); //THIS 1 FORCE DOWNLOAD
}


function generateTable(data , fn){
    var ws = XLSX.utils.json_to_sheet(data);
    console.log(ws);
    var temp = formatData(ws);
    var html_string = XLSX.utils.sheet_to_html(temp, { id: "data-table", editable: true });
    document.getElementById("showTable").innerHTML = html_string;
    document.getElementById("data-table").classList.add('.table-striped');
};



function formatData(obj) {
     var overall ={};   
    for (var key in obj) {
           
    if(key != "!ref"){

        var reA = /[^a-zA-Z]/; //use to remove digit
        var reN = /[^0-9]/;  // use to remove number
        //lower case 
        
        var cellDigit =parseInt(key.replace(reN,""));

        var cellAlpha = key.replace(reA,"");
      newkey =  cellAlpha +""+(cellDigit+4);
        overall[newkey] = obj[key];
      
    }else{
        overall[key]= "A1:S8";

    }
    }
    overall["A1"] ={t: "s", v: ">", w: ">"};
    overall["A2"] ={t: "s", v: ">", w: ">"};
    overall["A3"] ={t: "s", v: ">", w: ">"};
    overall["A4"] ={t: "s", v: ">", w: ">"};

    console.log(overall);
    return overall;
}
    //ajax call
function k2sReportGenerate(){
    $.ajax({ 
        url: 'http://47.88.132.185/api/v2/k2s_report', 
        headers: {  'X-DreamFactory-Api-Key': 'd5e2026a15103be156be2a2784bc7235756f05a61607a575338a49a58ec33cb8',
                    'Content-Type':'application/json'} ,
        method: "GET",
        success : function(result) { 
            var data = result.resource;
            var  num = 0;
            result =$.map(data, function (data) {
              
                
                    switch (data.transportMode){
                        case 1:
                        data.transportMode ="Ship";
                        break;
                        case 4:
                        data.transportMode ="Flight";
                        break;
                        default:
                        data.transportMode ="walk";

                    }

                return {
                    "No": ++num,
                    "Customs Reg No.":data.customsRegNo,
                    "Submission Date":data.createdAt,
                    "Destination":data.destination,
                    "Consignee":data.consignee,
                    "Consignor":data.consignor,
                    "Agent":data.agentName,
                    "Agent Code":data.agentCode,
                    "Invoice Amount":data.invoiceValue,
                    "Invoice Currency":data.invoiceCurrency,
                    "Exchange Rate":data.exchangeRate,
                    "Incoterm":data.incoTerm,
                    "Duty Amount":data.smkDutyAmount,
                    "Freight Amount":data.freightAmount,
                    "Freight Currency":data.freightCurrency,
                    "Exchange Rate":data.freightExchangeRate,
                    "Insurance Amount":data.insuranceAmount,
                    "Insurance Currency":data.insuranceCurrency,
                    "Exchange Rate":data.insuranceExchangeRate,
                    "Last Status":data.smkStatus,
                    "Transport Mode":data.transportMode
                }
            });
         
            generateTable(result);
        }, 
        error : function(result) { 
      //handle the error 
        console.log("i lose");
    } 
  }); }


  this.dataGen = function() {
    return k2sReportGenerate();
};

this.downloadFile= function(){return doit() };

}



ko.applyBindings(new reportViewModel());