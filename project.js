let request=require("request");

let ch=require("cheerio");
let fs=require("fs");
let path=require("path");
let xlsx=require("xlsx");

request("http://bpitindia.com/contact.html",getAllMatchUrl);
function getAllMatchUrl(err,res,html){
  //console.log(html);

  let STool=ch.load(html);
  let allProfile=STool("a");
  for(let i=0;i<allProfile.length/2;i++){
    if(STool(allProfile[i]).text()=="Faculty Profile"){
     let url=STool(allProfile[i]).attr("href");
     let FUrl="http://bpitindia.com/"+url;
     findDataOfranch(FUrl)
    }
  }
  //console.log(allProfile.length);
}

 function findDataOfranch(url){
   request(url,requestAns);
  function requestAns(err,res,html){
  
  let STool=ch.load(html);
  
  let AllNames=STool("div.col-md-4");
  let AllDetails=STool("div.col-md-8");
  
let branch=STool("h4.course-title").text();
let b=branch.split("Depatments");
branch=b[0].trim();
for(let i=0;i<AllNames.length;i++){

   let name=STool(AllNames[i]).find("a.d_inline.fw_600").text();
   let Table=STool(AllDetails[i]).find("table.table.table-bordered.table").find("tbody tr");
   let rCols=STool(Table).find("td");
   
   let Qual=STool(rCols[0]).text().trim();
    let email=STool(rCols[2]).text().trim();
    let exp=STool(rCols[4]).text().trim();
    let research=STool(rCols[6]).text().trim();
    let public=STool(rCols[8]).text().trim();
    let inter=STool(rCols[10]).text().trim();
    console.log(`Name: ${name} \nQualification: ${Qual} \nEmail: ${email} \nExperience:${exp} \nResearch: ${research} \nPublish: ${public} \nIntern: ${inter}`);
    console.log("///////////////////////////////////////////////////////////////////////////////////////////////////////////////");
   process(branch,name,Qual,email,exp,research,public,inter);
    
}
  }
  
  }

function process(branch,name,Qual,email,exp,research,public,inter){
   let dirPath=branch;
  let pStats = {
     branch:branch,
      Name:name,
      Qualification:Qual,
      Email:email,
      Experience:exp,
      Research:research,
      Publication:public,
      International_Publications:inter
  }

  if(fs.existsSync(dirPath)){//do nothing
    //console.log("Folder Exists");
}else{//create new folder
    fs.mkdirSync(dirPath);
}

let FilePath= path.join(dirPath, name+".xlsx");
let pData=[];
if(fs.existsSync(FilePath)){
pData = excelReader(FilePath, name);
pData.push(pStats);
}else{//i.e this is  first 
//create file
console.log("File ",FilePath,"created");
pData = [pStats];
}
excelWriter(FilePath,pData,name);

function excelReader(filePath, name){//got this function from stack overflow
  if(!fs.existsSync(filePath)){
      return null;
  }else{
      //workbook => excel
      let wt = xlsx.readFile(filePath);
      //get data from workbook
      let excelData = wt.Sheets[name];
      //convert excel format to json => array of object
      let ans = xlsx.utils.sheet_to_json(excelData);
      //console.log(ans);
      return ans;
  }
}

function excelWriter(filePath, json, name){
  //console.log(xlsx.readFile(filePath));
  let newWB = xlsx.utils.book_new();
  //console.log(json);
  let newWS = xlsx.utils.json_to_sheet(json);
  xlsx.utils.book_append_sheet(newWB, newWS, name);
  //file => create, replace
  xlsx.writeFile(newWB, filePath);
}
}