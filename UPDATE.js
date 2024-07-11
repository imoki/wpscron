// UPDATE.js 更新脚本
// 20240711

var confiWorkbook = 'CONFIG'  // 主配置表名称
var pushWorkbook = 'PUSH' // 推送表的名称
var emailWorkbook = 'EMAIL' // 邮箱表的名称

var workbook = [] // 存储已存在表数组

// 表中激活的区域的行数和列数
var row = 0;
var col = 0;
var maxRow = 100; // 规定最大行
var maxCol = 16; // 规定最大列
var colNum = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']

// 推送昵称(推送位置标识)选项：若“是”则推送“账户名称”，若账户名称为空则推送“单元格Ax”，这两种统称为位置标识。若“否”，则不推送位置标识
var configContent = [
    ['工作表的名称'],
    ['wps'],
  ]


// 定制化分配置表内容，WPS
var subConfigWps = [
  ['cookie(默认20个)', '文档名', '排除任务名', '时间范围'],
  ['xxxxxxxx1', '填这个文档的名称', 'CRON&PUSH', '0'],
]

// 定制化表
var subConfig = {
  "wps"  : subConfigWps,
}
// var mosaic = "xxxxxxxx" // 马赛克
// var strFail = "否"
// var strTrue = "是"

// 主函数执行流程

// createSubConfig()

function count2DArray(arr) {
  let count = 0;
  for (let i = 0; i < arr.length; i++) {
      for (let j = 0; j < arr[i].length; j++) {
          count++;
      }
  }
  return count;
}

let length = configContent.length - 1
// console.log(length)
console.log("🍳 正在检索分配表，并进行创建")
for (let i = 0; i < length; i++) {
  let workbook = ""
  let subworkbook = ""
  // console.log(configContent[i+1][4])
  if(1){  // 存活的才生成表
    // 部分表合并，如ddmc_ddgy，则以_为分割，生成ddmc表
    workbook = configContent[i+1][0].split("_")[0] // 使用下划线作为分隔符，取第一个部分)
    // ActivateSheet(workbook)  // 根据CONFIG表来生成
    // editConfigSheet(subConfigContent)   // editSubConfig()

    // 检查是否有定制化内容，有则生成
    try{
      subworkbook = subConfig[workbook]
      // console.log(subworkbook)
      if(subworkbook != undefined){
        ActivateSheet(workbook) // 激活阿里云盘分配置表
        editConfigSheet(subworkbook)  // editSubConfigCustomized(subConfigAliyundrive)
        // console.log("存在定制化内容")
      }
    }catch{
      // 无定制化内容
      // console.log("无定制化内容")
    }
  }
}

// // 加入跳转超链接
// hypeLink()

// CONFIG表添加超链接
function hypeLink(){
  let workSheet= Application.Sheets(confiWorkbook) //配置表
  let workUsedRowEnd = workSheet.UsedRange.RowEnd //用户使用表格的最后一行

  for(let row = 2; row <= workUsedRowEnd; row++){
    link_name=workSheet.Range("A" + row).Text
    if (link_name == "") {
      break; // 如果为空行，则提前结束读取
    }

    link_name ='=HYPERLINK("#'+link_name+'!$A$1","'+link_name+'")' //设置超链接
    //console.log(link_name)  // HYPERLINK("#PUSH!$A$1","PUSH")
    workSheet.Range("A" + row).Value =link_name 
  }
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

  console.log("✨ 当前激活表已存在：" + row + "行，" + col + "列")
}

// 激活工作表函数
function ActivateSheet(sheetName) {
  let flag = 0;
  try {
    // 激活工作表
    let sheet = Application.Sheets.Item(sheetName)
    sheet.Activate()
    console.log("🥚 激活工作表：" + sheet.Name)
    flag = 1;
  } catch {
    flag = 0;
    console.log("📢 无法激活工作表，工作表可能不存在")
    console.log("🎉 创建此表：" + sheetName)
    createSheet(sheetName)
  }
  return flag;
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

// 编辑主分配表(已弃用)
function editConfig() {
  determineRowCol();
  let lengthRow = configContent.length
  let lengthCol = configContent[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = configContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = configContent[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = configContent[i - 1][j]
    }
  }
}

// 编辑推送表(已弃用)
function editPush() {
  determineRowCol();
  let lengthRow = pushContent.length
  let lengthCol = pushContent[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = pushContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = pushContent[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = pushContent[i - 1][j]
    }
  }
}

// 编辑分配置表(已弃用)
function editSubConfig() {
  determineRowCol();
  let lengthRow = subConfigContent.length
  let lengthCol = subConfigContent[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = subConfigContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = subConfigContent[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = subConfigContent[i - 1][j]
    }
  }
}

// 编辑定制化分配置表(已弃用)
function editSubConfigCustomized(content) {
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


// 创建分配置表(已弃用)
function createSubConfig() {
  let length = subConfigWorkbook.length
  for (let i = 0; i < length; i++) {
    console.log("🥚 创建分配置表：" + subConfigWorkbook[i])
    createSheet(subConfigWorkbook[i])
  }
}

// 存储已存在的表
function storeWorkbook() {
  // 工作簿（Workbook）中所有工作表（Sheet）的集合,下面两种写法是一样的
  let sheets = Application.ActiveWorkbook.Sheets
  sheets = Application.Sheets

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