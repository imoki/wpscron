/*
    作者: imoki
    仓库: https://github.com/imoki/
    公众号：默库
    更新时间：20240714
    脚本：CRON_INIT.js 初始化程序，自动生成定时任务配置表。支持普通表格和智能表格。
    说明：要运行CRON.js之前，请先运行CRON_INIT脚本。
          并对表进行配置，第一次运行CRON_INIT会生成wps表，请先填写好wps表的内容。
          然后再运行一次CRON_INIT表生成CRON表，对CRON表进行配置。
          “是否调整”选项填“是”则会对其进行时间调整，默认为“否”是排除这个任务不会进行调整
          如果手动修改了定时任务时间，请重新运行一次CRON_INIT脚本，会自动生成最新的CRON配置表
*/

let sheetNameSubConfig = "wps"; // 分配置表名称
let sheetNameCron = "CRON"
var cookie = ""
var taskArray = []
var headers = ""
var count = "20" // 读取的文档页数
var excludeDocs = []
// 表中激活的区域的行数和列数
var row = 0;
var col = 0;
var maxRow = 100; // 规定最大行
var maxCol = 16; // 规定最大列
var workbook = [] // 存储已存在表数组
var colNum = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}

// 激活工作表函数
function ActivateSheet(sheetName) {
    let flag = 0;
    try {
      // 激活工作表
      let sheet = Application.Sheets.Item(sheetName);
      sheet.Activate();
      console.log("🥚 激活工作表：" + sheet.Name);
      flag = 1;
    } catch {
      flag = 0;
      console.log("🍳 无法激活工作表，工作表可能不存在");
    }
    return flag;
}

// 存储已存在的表
function storeWorkbook() {
  // 工作簿（Workbook）中所有工作表（Sheet）的集合,下面两种写法是一样的
  let sheets = Application.ActiveWorkbook.Sheets
  sheets = Application.Sheets

  // 清空数组
  workbook.length = 0

  // 打印所有工作表的名称
  for (let i = 1; i <= sheets.Count; i++) {
    workbook[i - 1] = (sheets.Item(i).Name)
    // console.log(workbook[i-1])
  }
}

// 判断表是否已存在
function workbookComp(name) {
  let flag = 0;
  let length = workbook.length
  for (let i = 0; i < length; i++) {
    if (workbook[i] == name) {
      flag = 1;
      console.log("✨ " + name + "表已存在")
      break
    }
  }
  return flag
}

// 创建表，若表已存在则不创建，直接写入数据
function createSheet(name) {
  // const defaultName = Application.Sheets.DefaultNewSheetName
  // 工作表对象
  if (!workbookComp(name)) {
    Application.Sheets.Add(
      null,
      Application.ActiveSheet.Name,
      1,
      Application.Enum.XlSheetType.xlWorksheet,
      name
    )
  }
}


// 获取wps_sid、cookie
function getWpsSid(){
  // flagConfig = ActivateSheet(sheetNameSubConfig); // 激活wps配置表
  // 主配置工作表存在
  if (1) {
    console.log("🍳 开始读取wps配置表");
    for (let i = 2; i <= 100; i++) {
      // 读取wps表格配置
      wps_sid = Application.Range("A" + i).Text; // 以第一个wps为准
      // name = Application.Range("H" + i).Text;
      
      excludeDocs = Application.Range("C" + i).Text.split("&")

      break
    }
  }
  return wps_sid
  
  // filename = name
}

// 获取file_id
function getFile(url){
  // 查看定时任务
  resp = HTTP.get(
    url,
    { headers: headers }
  );

  resp = resp.json()
  // console.log(resp)
  resplist = resp["list"]
  let cronlist = ""
  for(let i = 0; i<resplist.length; i++){
    roaming = resplist[i]["roaming"]
    // console.log(roaming)
    fileid = roaming["fileid"]
    name = roaming["name"]
    if(juiceXLSX(name)){
      // console.log(name.split(".")[0])
      if(juiceDocs(name.split(".")[0])){
        console.log("🏹 排除 " + name + " 文档")
      }else{
        console.log("🎯 存在 " + name + " 文档")
        cronlist = taskExist(fileid)
        if(cronlist.length > 0){
          console.log("🎉 存在定时任务")
          // console.log(cronlist)
          for(let i = 0; i < cronlist.length; i++){
            
            task = cronlist[i]
            task_id = task["task_id"]
            script_id = task["script_id"]
            script_name = task["script_name"]

            cron_detail = task["cron_detail"]
            cron_desc = cron_detail["cron_desc"]
            cron_type = cron_desc["cron_type"]
            day_of_month = cron_desc["day_of_month"]
            day_of_week = cron_desc["day_of_week"]
            // month = cron_desc["month"]
            hour = cron_desc["hour"]
            minute = cron_desc["minute"]
            // year = cron_desc["year"]
            // file_id = fileid
            taskArray.push({
              "filename" : name,
              "fileid" : fileid,
              "script_id" : script_id,
              "script_name" : script_name,
              "task_id" : task_id,
              "cron_type":cron_type,
              "day_of_month": day_of_month,
              "day_of_week": day_of_week,
              "hour"  : hour,
              "minute" : minute,
            })

          }
        }

      }
      


      // console.log("🍳 file_id : " + file_id)
      // break
    }
  }

  // console.log(taskArray)
  sleep(3000)
}

// 判断是否为xlsx文件
function juiceXLSX(name){
  let flag = 0
  let array= name.split(".") // 使用|作为分隔符
  if(array.length == 2 && (array[1] == "xlsx" || array[1] == "ksheet")){
    flag = 1
  }
  return flag 
}

// 判断是否为要排除文件
function juiceDocs(name){
  let flag = 0
  if((excludeDocs.length == 1 && excludeDocs[0] == "") || excludeDocs.length == 0){
    flag = 0
    // console.log("excludeDocs不符合")
  }else{
    for(let i= 0; i<excludeDocs.length; i++){
      if(name == excludeDocs[i]){
        flag = 1  // 找到要排除的文档了
        // console.log("找到要排除的文档了")
      }
    }
  }
  
  return flag 
}

// 判断是否存在定时任务
function taskExist(file_id){
  url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/cron_tasks";
  // console.log(url)
  // 查看定时任务
  resp = HTTP.get(
    url,
    { headers: headers }
  );

  resp = resp.json()
  // console.log(resp)
  // list -> 数组 -> file_id、task_id、script_name，cron_detail->字典
  cronlist = resp["list"]
  sleep(3000)
  return cronlist
}

function writeTask(){
  createSheet(sheetNameCron)
  flagConfig = ActivateSheet(sheetNameCron); // 激活cron配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    // console.log(taskArray.length)
    console.log("🍳 开始读取CRON配置表");
    let pos = 1
    // 写入表头
    Application.Range(colNum[0] + pos).Value = "文档名"
    Application.Range(colNum[1] + pos).Value = "文档id"
    Application.Range(colNum[2] + pos).Value = "任务名"
    Application.Range(colNum[3] + pos).Value = "任务id"
    Application.Range(colNum[4] + pos).Value = "脚本id"
    Application.Range(colNum[5] + pos).Value = "是否调整"
    Application.Range(colNum[6] + pos).Value = "时间范围"
    Application.Range(colNum[7] + pos).Value = "当前定时时间"
    Application.Range(colNum[8] + pos).Value = "额外参数"
    // taskArray.push({
    //     "filename" : name,
    //     "fileid" : fileid,
    //     "script_id" : script_id,
    //     "script_name" : script_name,
    //     "task_id" : task_id,
    //     "cron_type":cron_type,
    //     "day_of_month": day_of_month,
    //     "day_of_week": day_of_week,
    //     "hour"  : hour,
    //     "minute" : minute,
    //   })

    // console.log(taskArray)
    // 写入获取到的定时任务数据
    for(let i = 0; i < taskArray.length; i++){
      pos = i + 2
      let j = 0
      Application.Range(colNum[0] + pos).Value = taskArray[i]["filename"]
      Application.Range(colNum[1] + pos).Value = taskArray[i]["fileid"]
      Application.Range(colNum[2] + pos).Value = taskArray[i]["script_name"]
      Application.Range(colNum[3] + pos).Value = taskArray[i]["task_id"]
      Application.Range(colNum[4] + pos).Value = taskArray[i]["script_id"]
      Application.Range(colNum[5] + pos).Value = "否"
      Application.Range(colNum[6] + pos).Value = "0~23"
      Application.Range(colNum[7] + pos).Value = taskArray[i]["hour"] + ":" + taskArray[i]["minute"]
      Application.Range(colNum[8] + pos).Value = taskArray[i]["cron_type"] + "&" + taskArray[i]["day_of_month"] + "&" + taskArray[i]["day_of_week"]
    }
  }
}

function initCron(){
  try{
    Application.Sheets.Item('CRON').Delete()  // 为了获得最新数据，删除CRON表
    storeWorkbook()
  }catch{
    console.log("🍳 不存在CRON表，开始进行创建")
  }
  
  // 获取file_id
  url = "https://drive.kdocs.cn/api/v5/roaming?count=" + count  // 只对前20条进行判断
  getFile(url)
  writeTask()

  console.log("✨ 已完成对CRON表的写入，请到CRON表进行配置")
  console.log("✨ 然后将CRON.js脚本加入定时任务，即可自动调整定时任务时间")
}

// 判断表格行列数，并记录目前已写入的表格行列数。目的是为了不覆盖原有数据，便于更新
function determineRowCol() {
  for (let i = 1; i < maxRow; i++) {
    let content = Application.Range("A" + i).Text
    if (content == "")  // 如果为空行，则提前结束读取
    {
      row = i - 1;  // 记录的是存在数据所在的行
      break;
    }
  }
  // 超过最大行了，认为row为0，从头开始
  let length = colNum.length
  for (let i = 1; i <= length; i++) {
    content = Application.Range(colNum[i - 1] + "1").Text
    if (content == "")  // 如果为空行，则提前结束读取
    {
      col = i - 1;  // 记录的是存在数据所在的行
      break;
    }
  }
  // 超过最大行了，认为col为0，从头开始

  // console.log("✨ 当前激活表已存在：" + row + "行，" + col + "列")
}

// 统一编辑表函数
function editConfigSheet(content) {
  determineRowCol();
  let lengthRow = content.length
  let lengthCol = content[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = content[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = content[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = content[i - 1][j]
    }
  }
}

// 创建wps表
function createWpsConfig(){
  createSheet(sheetNameSubConfig) // 若wsp表不存在创建wps表
  let flagExitContent = 1

  if(ActivateSheet(sheetNameSubConfig)) // 激活cron配置表
  {
    // wps表内容
    let content = [
      ['wps_sid', '任务配置表超链接', '排除文档'],
      ['此处填写wps_sid', '点击此处跳转到CRON表', '']
    ]
    determineRowCol() // 读取函数
    if(row <= 1 || col < content[0].length){ // 说明是空表或只有表头未填写内容，或者表格有新增列内容则需要先填写
      // console.log(row)
      flagExitContent = 0 // 原先不存在内容，告诉用户先填内容
      editConfigSheet(content)
      // console.log(row)
      let name = "点击此处跳转到CRON表"  // 'CRON'!A1
      let link = "CRON"
      let link_name ='=HYPERLINK("#'+link+'!$A$1","'+name+'")' //设置超链接
      //console.log(link_name)  // HYPERLINK("#PUSH!$A$1","PUSH")
      Application.Range("B2").Value = link_name
    }
  }

  return flagExitContent
  
}


function main(){
  storeWorkbook()
  let flagExitContent = createWpsConfig()
  if(flagExitContent == 0){
    console.log("📢 请先填写wps表，然后再运行一次此脚本")
  }else{
    wps_sid = getWpsSid() // 获取wps_sid
    cookie = "wps_sid=" + wps_sid // 获取cookie
    // console.log(excludeDocs)

    headers = {
      "Cookie": cookie,
      "Content-Type" : "application/json",
      "Origin":"https://www.kdocs.cn",
      "Priority":"u=1, i",
    }
    
    
    // 获取定时任务,生成CRON定时任务表
    initCron()

  }

}

main()