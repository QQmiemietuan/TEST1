$(function() {
    $(".header-page").load("../01header.html");
    $(".header2-page").load("../02header.html");
    $(".nav-page").load("../03nav.html");
    $(".line-page").load("../04line.html");
    $(".line1-page").load("../04_1line.html");
    $(".line2-page").load("../04_2line.html");
    $(".line3-page").load("../04_3line.html");
    $(".line4-page").load("../04_4line.html");
    $(".line5-page").load("../04_5line.html");
    $(".line6-page").load("../04_6line.html");
    $(".line7-page").load("../04_7line.html");
    $(".nav2-page").load("../05nav2.html");
    $(".footer-page").load("../footer.html");
});
function search1(){
    var a="";
    btype = document.getElementById("btype").value;
    inMin = document.getElementById("inMin").value;
    inMax = document.getElementById("inMax").value;
    Wmin = document.getElementById("Wmin").value;
    Wmax = document.getElementById("Wmax").value;
    outmin = document.getElementById("outmin").value;
    outmax = document.getElementById("outmax").value;
    if(typeof btype =="null")
        a=a+"軸承型號無輸入";
    else
        a=a+"軸承型號:"+btype;
    if(typeof inMin =="null")
        a=a+"最小內徑無輸入";
    else
        a=a+";"+"<br>內徑"+inMin;
    if(typeof inMax =="null")
        a=a+"最大內徑無輸入";
    else
        a=a+"<"+inMax;
    if(typeof Wmin =="null")
        a=a+"最小寬度無輸入";
    else
        a=a+";"+"<br>寬度"+Wmin;
    if(typeof Wmax =="null")
        a=a+"最大寬度無輸入";
    else
        a=a+"<"+Wmax;
    if(typeof outmin =="null")
        a=a+"最小外徑無輸入";
    else
        a=a+";"+"<br>外徑"+outmin;
    if(typeof outmax =="null")
        a=a+"最大外徑無輸入";
    else
        a=a+"<"+outmax;
    if (window.localStorage) { 
        //儲存變數的值 
        localStorage.name = a; 
        location.href = '02-1-1.html'; 
    } else { 
        alert("NOT SUPPORT"); 
    } 
}
function clear1(){
}
function getvalue(){
    var value = localStorage["name"];
    alert("value="+value);
    document.getElementById("content").value = value;
    if(document.getElementById("content").value =="null")
      alert("value!="+value);
      readWorkbookFromLocalFile('../txt/txt.xlsx',value);  
}
function readWorkbookFromLocalFile(file, value3) {
    alert("獨到囉");
	var workbook = new Excel.Workbook(); 
    workbook.xlsx.readFile('txt/txt.xlsx')
    .then(function() {
        var worksheet = workbook.getWorksheet(sheet);
        worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
          console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
        });
    });
    document.getElementById("content1").value = workbook;

}