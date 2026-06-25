import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ─── THEME ───────────────────────────────────────────────────────────────────
const C = {
  sidebar:"#1A2340", sidebarA:"rgba(59,130,246,0.22)", sidebarT:"rgba(255,255,255,0.6)",
  accent:"#2563EB", accentL:"#EFF6FF", accentT:"#1D4ED8",
  bg:"#F0F2F5", card:"#FFFFFF", text:"#0F172A", muted:"#64748B", border:"#E2E8F0",
  success:"#059669", successL:"#D1FAE5", successD:"#065F46",
  warning:"#D97706", warningL:"#FEF3C7", warningD:"#92400E",
  danger:"#DC2626", dangerL:"#FEE2E2", dangerD:"#991B1B",
  purple:"#7C3AED", purpleL:"#EDE9FE",
  orange:"#EA580C", orangeL:"#FFEDD5",
  teal:"#0D9488", tealL:"#CCFBF1",
};

// ─── BILINGUAL TRANSLATIONS ───────────────────────────────────────────────────
const TX = {
  en:{
    // Nav
    dashboard:"Dashboard", attendance:"Attendance", leaves:"Leave", claims:"Claims",
    payslip:"Payslip", memo:"HR Memo", policy:"Company Policy",
    training:"Training", feedback:"Feedback", appraisal:"Appraisal",
    calendar:"Calendar", shift:"Shift Schedule", settings:"Settings",
    employees:"Employees", payroll:"Payroll", reports:"Reports",
    // Common
    add:"Add", edit:"Edit", delete:"Delete", save:"Save", cancel:"Cancel",
    submit:"Submit", approve:"Approve", reject:"Reject", view:"View",
    search:"Search", filter:"Filter", export:"Export Excel", close:"Close",
    confirm:"Confirm", yes:"Yes", no:"No", back:"Back",
    name:"Name", date:"Date", status:"Status", action:"Action",
    from:"From", to:"To", days:"Days", reason:"Reason", amount:"Amount",
    type:"Type", description:"Description", remarks:"Remarks",
    // Attendance
    clockIn:"Clock In", clockOut:"Clock Out", photo:"Take Photo",
    gps:"GPS Location", hoursWorked:"Hours Worked",
    site:"Work Site", capturing:"Capturing location...",
    locationCaptured:"Location captured", usePhoto:"Use Photo",
    // Leave
    leaveBalance:"Leave Balance", applyLeave:"Apply Leave",
    entitlement:"Entitlement", used:"Used", balance:"Balance",
    pendingSup:"Pending Supervisor", pendingHR:"Pending HR",
    // Claims
    submitClaim:"Submit Claim", claimType:"Claim Type",
    receipt:"Attach Receipt", ot:"Overtime", transport:"Transport",
    medical:"Medical", meal:"Meal", others:"Others",
    // Payslip
    basicSalary:"Basic Salary", allowance:"Allowance", grossPay:"Gross Pay",
    cpfEmployee:"CPF (Employee)", cpfEmployer:"CPF (Employer)",
    netPay:"Net Pay", sdl:"SDL", fwl:"FWL", otPay:"OT Pay",
    // Memo
    newMemo:"New Announcement", viewMemo:"View Announcement",
    // Policy
    uploadPolicy:"Upload Policy", effectiveDate:"Effective Date",
    // Training
    venue:"Venue", trainer:"Trainer", duration:"Duration",
    // Feedback
    anonymous:"Submit Anonymously", category:"Category",
    // Appraisal
    selfAppraisal:"Self Appraisal", rating:"Rating",
    supervisorReview:"Supervisor Review",
    // Shift
    shift:"Shift", dayShift:"Day", nightShift:"Night", off:"Off",
    // Admin
    signOut:"Sign Out", language:"Language", processed:"Processed",
    draft:"Draft", published:"Published",
    pendingApproval:"Pending Approval",
    totalEmployees:"Total Employees",
    active:"Active",
    // Status
    Approved:"Approved", Rejected:"Rejected", Pending:"Pending",
    "Pending Supervisor":"Pending Supervisor", "Pending HR":"Pending HR",
  },
  zh:{
    dashboard:"仪表板", attendance:"出勤", leaves:"请假", claims:"报销",
    payslip:"工资单", memo:"人事通告", policy:"公司政策",
    training:"培训", feedback:"反馈", appraisal:"绩效考核",
    calendar:"日历", shift:"排班", settings:"设置",
    employees:"员工", payroll:"薪资", reports:"报告",
    add:"添加", edit:"编辑", delete:"删除", save:"保存", cancel:"取消",
    submit:"提交", approve:"批准", reject:"拒绝", view:"查看",
    search:"搜索", filter:"筛选", export:"导出Excel", close:"关闭",
    confirm:"确认", yes:"是", no:"否", back:"返回",
    name:"姓名", date:"日期", status:"状态", action:"操作",
    from:"从", to:"至", days:"天数", reason:"原因", amount:"金额",
    type:"类型", description:"描述", remarks:"备注",
    clockIn:"打卡上班", clockOut:"打卡下班", photo:"拍照",
    gps:"GPS位置", hoursWorked:"工作时间",
    site:"工地", capturing:"正在获取位置...",
    locationCaptured:"位置已获取", usePhoto:"使用照片",
    leaveBalance:"假期余额", applyLeave:"申请休假",
    entitlement:"应享天数", used:"已用天数", balance:"余额",
    pendingSup:"待主管审批", pendingHR:"待人事审批",
    submitClaim:"提交报销", claimType:"报销类型",
    receipt:"附上收据", ot:"加班费", transport:"交通费",
    medical:"医疗费", meal:"餐费", others:"其他",
    basicSalary:"基本工资", allowance:"津贴", grossPay:"总收入",
    cpfEmployee:"公积金（员工）", cpfEmployer:"公积金（雇主）",
    netPay:"实发工资", sdl:"技能发展税", fwl:"外劳税", otPay:"加班工资",
    newMemo:"新通告", viewMemo:"查看通告",
    uploadPolicy:"上传政策", effectiveDate:"生效日期",
    venue:"地点", trainer:"培训师", duration:"时长",
    anonymous:"匿名提交", category:"类别",
    selfAppraisal:"自我评估", rating:"评分",
    supervisorReview:"主管评审",
    shift:"班次", dayShift:"白班", nightShift:"夜班", off:"休息",
    signOut:"退出登录", language:"语言", processed:"已处理",
    draft:"草稿", published:"已发布",
    pendingApproval:"待审批",
    totalEmployees:"员工总数",
    active:"在职",
    Approved:"已批准", Rejected:"已拒绝", Pending:"待处理",
    "Pending Supervisor":"待主管审批", "Pending HR":"待人事审批",
  }
};
const t=(lang,key)=>TX[lang]?.[key]||TX.en?.[key]||key;

// ─── INITIAL DATA ─────────────────────────────────────────────────────────────
const SUPER_ADMIN = {
  id:"sa", username:"info@custera.com.sg", password:"Custera2025!",
  role:"superadmin", name:"Super Admin", empId:"", dept:"", reportingOfficerId:"", active:true
};

const INIT_USERS = [
  SUPER_ADMIN,
  { id:"u1", username:"admin", password:"admin123", role:"admin", name:"HR Admin (Clara)", empId:"C010", dept:"HR & Admin.", reportingOfficerId:"", active:true },
  { id:"u2", username:"zou", password:"zou123", role:"supervisor", name:"ZOU LI", empId:"C006", dept:"HR & Admin.", reportingOfficerId:"u1", active:true },
  { id:"u3", username:"willie", password:"willie123", role:"employee", name:"ANG Hwee Lee, Willie", empId:"C001", dept:"Management", reportingOfficerId:"u1", active:true },
  { id:"u4", username:"ge", password:"ge123", role:"employee", name:"GE Qiu Zhang", empId:"C002", dept:"Engineering & Operations", reportingOfficerId:"u2", active:true },
  { id:"u5", username:"tang", password:"tang123", role:"employee", name:"TANG Chin Yeon", empId:"C004", dept:"Business Development", reportingOfficerId:"u2", active:true },
];

const INIT_EMPLOYEES = [
  { id:"C001",name:"ANG Hwee Lee, Willie",fin:"S7340234D",company:"STEC",position:"CEO",department:"Management",dateJoined:"2011-05-01",gender:"M",dob:"1973-11-05",nationality:"Singaporean",workPass:"SC",epExpiry:"",passportNo:"K4580555H",passportExpiry:"2033-12-24",mobile:"9688 2560",address:"321 Choa Chu Kang Ave 3 #11-21",postal:"689864",marital:"Divorced",basicSalary:0,allowance:0,bankName:"OCBC",bankAccount:"6952 3635 6001",annualLeave:14,status:"Active",site:"HQ" },
  { id:"C002",name:"GE Qiu Zhang",fin:"M4624560U",company:"Custera",position:"Chief Engineering Officer",department:"Engineering & Operations",dateJoined:"2026-01-27",gender:"M",dob:"1982-10-20",nationality:"Chinese",workPass:"EP",epExpiry:"2028-01-30",passportNo:"PE3315687",passportExpiry:"2030-03-06",mobile:"8982 2219",address:"Blk 80 Compassvale Bow #08-40",postal:"544570",marital:"",basicSalary:8400,allowance:2200,bankName:"DBS",bankAccount:"2729124560",annualLeave:14,status:"Active",site:"Tuas Site" },
  { id:"C004",name:"TANG Chin Yeon",fin:"S9779502I",company:"UTEC",position:"BD Engineer",department:"Business Development",dateJoined:"2022-02-01",gender:"M",dob:"1997-07-08",nationality:"Malaysian",workPass:"SPR",epExpiry:"",passportNo:"K54983963",passportExpiry:"2031-12-06",mobile:"8542 5468",address:"APT BLK 280B Sengkang East Ave #15-629",postal:"542280",marital:"Single",basicSalary:4500,allowance:500,bankName:"DBS",bankAccount:"2710614542",annualLeave:14,status:"Active",site:"Jurong Site" },
  { id:"C005",name:"CHUA Cheng Yi",fin:"S9686151F",company:"UTEC",position:"Site Engineer",department:"Engineering & Operations",dateJoined:"2022-04-01",gender:"M",dob:"1996-07-28",nationality:"Malaysian",workPass:"SPR",epExpiry:"",passportNo:"A55417959",passportExpiry:"2031-08-20",mobile:"9340 2639",address:"Blk 13 Cantonment Close #26-29",postal:"080013",marital:"Single",basicSalary:4200,allowance:400,bankName:"DBS",bankAccount:"2711397501",annualLeave:14,status:"Active",site:"Tuas Site" },
  { id:"C006",name:"ZOU LI",fin:"S7579864D",company:"STEC",position:"Asst Admin Manager",department:"HR & Admin.",dateJoined:"2025-10-15",gender:"F",dob:"1975-05-28",nationality:"Singaporean",workPass:"SC",epExpiry:"",passportNo:"",passportExpiry:"",mobile:"9748 8442",address:"53 Serangoon Terrace",postal:"535787",marital:"Married",basicSalary:4800,allowance:100,bankName:"OCBC",bankAccount:"",annualLeave:14,status:"Active",site:"HQ" },
  { id:"C010",name:"YANG Zi Qing, Clara",fin:"T0112349H",company:"STEC",position:"HR cum Admin Exec",department:"HR & Admin.",dateJoined:"2026-01-19",gender:"F",dob:"2001-03-28",nationality:"Singaporean",workPass:"SC",epExpiry:"",passportNo:"K4382218H",passportExpiry:"2033-09-10",mobile:"9053 2999",address:"201 Marsiling Drive #07-112",postal:"730201",marital:"Single",basicSalary:3500,allowance:50,bankName:"UOB",bankAccount:"453-300-635-8",annualLeave:14,status:"Active",site:"HQ" },
  { id:"C011",name:"LAU Ah Kwai",fin:"S1218010A",company:"STEC",position:"Pantry Assistant",department:"Admin",dateJoined:"2026-03-19",gender:"F",dob:"1956-01-08",nationality:"Singaporean",workPass:"SC",epExpiry:"",passportNo:"",passportExpiry:"",mobile:"88031570",address:"Blk 320 Jurong East St #07-72",postal:"600320",marital:"Married",basicSalary:1900,allowance:0,bankName:"POSB",bankAccount:"151-07403-0",annualLeave:7,status:"Active",site:"HQ" },
  { id:"C012",name:"TEE Siang Long",fin:"S8268798Z",company:"STEC",position:"Quantity Surveyor",department:"Commercial",dateJoined:"2026-05-01",gender:"M",dob:"1982-06-10",nationality:"Malaysian",workPass:"SPR",epExpiry:"",passportNo:"",passportExpiry:"",mobile:"9027 9197",address:"357A Admiralty Drive",postal:"751357",marital:"Married",basicSalary:5000,allowance:50,bankName:"",bankAccount:"",annualLeave:14,status:"Active",site:"Jurong Site" },
];

const SG_LEAVE_TYPES = [
  {name:"Annual Leave",days:14,paid:true},{name:"Medical Leave",days:14,paid:true},
  {name:"Hospitalisation Leave",days:60,paid:true},{name:"Childcare Leave",days:6,paid:true},
  {name:"Maternity Leave",days:112,paid:true},{name:"Paternity Leave",days:14,paid:true},
  {name:"Compassionate Leave",days:3,paid:true},{name:"No Pay Leave",days:999,paid:false},
  {name:"Marriage Leave",days:3,paid:true},
];

const INIT_LEAVES = [
  {id:1,empId:"C004",type:"Annual Leave",from:"2026-07-01",to:"2026-07-03",days:3,reason:"Personal",status:"Pending Supervisor",supervisorComment:"",hrComment:"",submittedDate:"2026-06-20"},
  {id:2,empId:"C006",type:"Medical Leave",from:"2026-06-18",to:"2026-06-18",days:1,reason:"MC from clinic",status:"Approved",supervisorComment:"Approved",hrComment:"Noted",submittedDate:"2026-06-18"},
];

const INIT_CLAIMS = [
  {id:1,empId:"C004",type:"Transport",date:"2026-06-15",amount:12.50,description:"Taxi to Tuas site",status:"Pending Supervisor",supervisorComment:"",hrComment:"",receipt:"",submittedDate:"2026-06-15"},
  {id:2,empId:"C005",type:"OT",date:"2026-06-14",amount:180,description:"OT 4hrs - concrete pour",status:"Pending HR",supervisorComment:"Verified on site",hrComment:"",receipt:"",submittedDate:"2026-06-14"},
];

const INIT_MEMOS = [
  {id:1,title:"Safety Briefing - Tuas Site",titleZh:"安全简报 - 大士工地",body:"All workers at Tuas site must attend the mandatory safety briefing on 28 June 2026 at 7:30am. PPE must be worn at all times on site.",bodyZh:"所有大士工地工人必须于2026年6月28日上午7时30分参加强制性安全简报。工地上必须全程佩戴个人防护装备。",date:"2026-06-20",tag:"Safety",author:"HR Admin",targetDepts:[],readBy:[]},
  {id:2,title:"Public Holiday - Hari Raya Haji",titleZh:"公共假日 - 哈芝节",body:"The office and all sites will be closed on 6 June 2026 for Hari Raya Haji.",bodyZh:"2026年6月6日（哈芝节）办公室及所有工地将休息一天。",date:"2026-05-28",tag:"Holiday",author:"HR Admin",targetDepts:[],readBy:[]},
];

const INIT_POLICIES = [
  {id:1,title:"Site Safety Policy",titleZh:"工地安全政策",category:"Safety",version:"v2.1",effectiveDate:"2026-01-01",content:"1. All workers must wear PPE at all times on construction sites.\n2. Toolbox meetings mandatory every Monday morning.\n3. Near-miss incidents must be reported within 24 hours.\n4. Working at height requires valid WAH permit.\n5. No mobile phone use while operating machinery.",contentZh:"1. 所有工人在建筑工地必须全程佩戴个人防护装备。\n2. 每周一上午必须举行工具箱会议。\n3. 险肇事件必须在24小时内上报。\n4. 高空作业需持有效WAH许可证。\n5. 操作机械时禁止使用手机。",uploadedBy:"HR Admin",date:"2026-01-01"},
  {id:2,title:"Leave Policy",titleZh:"请假政策",category:"HR",version:"v3.0",effectiveDate:"2026-01-01",content:"1. Annual leave must be applied 3 days in advance.\n2. Medical leave requires MC from registered clinic.\n3. Emergency leave may be applied same day with supervisor approval.\n4. No-pay leave requires management approval.",contentZh:"1. 年假须提前3天申请。\n2. 病假需提供注册诊所的医疗证明。\n3. 紧急假可当天申请，需主管批准。\n4. 无薪假需管理层批准。",uploadedBy:"HR Admin",date:"2026-01-01"},
];

const INIT_TRAININGS = [
  {id:1,title:"Working at Height (WAH) Certification",titleZh:"高空作业认证",description:"Mandatory certification for all site workers working above 2m.",descriptionZh:"所有在2米以上工作的工地工人的强制认证。",venue:"BCA Academy, Buona Vista",trainer:"BCA Certified Trainer",date:"2026-07-15",time:"08:00",duration:"8 hours",targetDepts:["Engineering & Operations"],assignedTo:["C004","C005"],status:"Upcoming"},
  {id:2,title:"HR System Orientation",titleZh:"人事系统培训",description:"Introduction to the new Custera.HR system for all staff.",descriptionZh:"向所有员工介绍新的Custera.HR系统。",venue:"HQ Meeting Room",trainer:"HR Admin",date:"2026-07-01",time:"14:00",duration:"2 hours",targetDepts:[],assignedTo:[],status:"Upcoming"},
];

const INIT_FEEDBACK_SECTIONS = [
  {id:1,name:"Safety Concerns",nameZh:"安全问题",description:"Report any safety hazards or concerns on site",active:true},
  {id:2,name:"Workplace Improvement",nameZh:"工作环境改善",description:"Suggestions for improving the work environment",active:true},
  {id:3,name:"Management Feedback",nameZh:"管理层反馈",description:"Feedback on management practices",active:true},
];

const INIT_FEEDBACKS = [
  {id:1,empId:"C005",sectionId:1,subject:"Scaffolding at Block C",subjectZh:"",message:"The scaffolding at Block C level 4 looks unstable. Needs inspection.",isAnonymous:false,date:"2026-06-18",status:"Open",adminReply:""},
];

const INIT_APPRAISAL_FORMS = [
  {id:1,title:"H1 2026 Performance Review",titleZh:"2026年上半年绩效考核",period:"Jan–Jun 2026",deadline:"2026-07-31",targetDepts:[],status:"Active",
    questions:[
      {id:"q1",question:"Work Quality",questionZh:"工作质量",type:"rating",category:"Performance"},
      {id:"q2",question:"Safety Awareness",questionZh:"安全意识",type:"rating",category:"Safety"},
      {id:"q3",question:"Teamwork",questionZh:"团队合作",type:"rating",category:"Behaviour"},
      {id:"q4",question:"Attendance & Punctuality",questionZh:"出勤与守时",type:"rating",category:"Attendance"},
      {id:"q5",question:"What areas do you feel you have improved?",questionZh:"您认为自己在哪些方面有所进步？",type:"text",category:"Development"},
      {id:"q6",question:"What support do you need from management?",questionZh:"您需要管理层提供什么支持？",type:"text",category:"Development"},
    ]
  },
];

const INIT_APPRAISAL_SUBMISSIONS = [];

const INIT_ATTENDANCE = [
  {id:1,empId:"C004",date:"2026-06-25",clockIn:"07:45",clockOut:"17:30",lat:1.3521,lng:103.8198,site:"Tuas Site",photo:null,status:"Present",hoursWorked:9.75},
  {id:2,empId:"C005",date:"2026-06-25",clockIn:"07:50",clockOut:"17:45",lat:1.3521,lng:103.8198,site:"Tuas Site",photo:null,status:"Present",hoursWorked:9.92},
];

const INIT_SHIFTS = [
  {id:1,empId:"C004",weekStart:"2026-06-23",shifts:[
    {date:"2026-06-23",type:"Day",start:"07:00",end:"17:00",site:"Jurong Site"},
    {date:"2026-06-24",type:"Day",start:"07:00",end:"17:00",site:"Jurong Site"},
    {date:"2026-06-25",type:"Day",start:"07:00",end:"17:00",site:"Jurong Site"},
    {date:"2026-06-26",type:"Day",start:"07:00",end:"17:00",site:"Jurong Site"},
    {date:"2026-06-27",type:"Day",start:"07:00",end:"17:00",site:"Jurong Site"},
    {date:"2026-06-28",type:"Off",start:"",end:"",site:""},
    {date:"2026-06-29",type:"Off",start:"",end:"",site:""},
  ]},
];

const INIT_CALENDAR_EVENTS = [
  {id:1,title:"Public Holiday - Hari Raya Haji",type:"Holiday",date:"2026-06-06",endDate:"2026-06-06",dept:"All",color:C.warning,createdBy:"HR Admin"},
  {id:2,title:"Safety Briefing - Tuas",type:"Meeting",date:"2026-06-28",endDate:"2026-06-28",dept:"Engineering & Operations",color:C.danger,createdBy:"HR Admin"},
];

const INIT_PAYROLL = [];

// ─── HELPERS ──────────────────────────────────────────────────────────────────
let _uid = 700;
function uid(){return ++_uid;}
function todayStr(){return new Date().toISOString().slice(0,10);}
function fmtDate(d){if(!d)return"-";try{return new Date(d).toLocaleDateString("en-SG",{day:"2-digit",month:"short",year:"numeric"});}catch(e){return d;}}
function daysUntil(d){if(!d)return null;return Math.ceil((new Date(d)-new Date())/86400000);}
function calcAge(dob){if(!dob)return"";return Math.floor((Date.now()-new Date(dob).getTime())/3.154e10);}
function sgd(n){return"S$"+Number(n||0).toLocaleString("en-SG",{minimumFractionDigits:2,maximumFractionDigits:2});}
function calcCPF(basic,nat){
  if(nat==="Singaporean"||nat==="SPR"){return{employee:Math.round(Math.min(basic*0.20,1200)),employer:Math.round(Math.min(basic*0.17,1020))};}
  return{employee:0,employer:0};
}
function calcSDF(basic){return Math.max(2,Math.ceil(basic*0.0025));}
function genPwd(){const c="ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!@#";let p="";for(let i=0;i<10;i++)p+=c[Math.floor(Math.random()*c.length)];return p;}
function exportXLS(data,sheet,file){if(!data?.length){alert("No data");return;}const ws=XLSX.utils.json_to_sheet(data);ws["!cols"]=Object.keys(data[0]).map(k=>({wch:Math.min(40,Math.max(k.length+2,...data.map(r=>String(r[k]||"").length)))}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheet.slice(0,31));XLSX.writeFile(wb,file);}
function exportMulti(sheets,file){const wb=XLSX.utils.book_new();sheets.forEach(({n,data})=>{if(!data?.length)return;const ws=XLSX.utils.json_to_sheet(data);ws["!cols"]=Object.keys(data[0]).map(k=>({wch:Math.min(40,Math.max(k.length+2,...data.map(r=>String(r[k]||"").length),10))}));XLSX.utils.book_append_sheet(wb,ws,n.slice(0,31));});XLSX.writeFile(wb,file);}

// ─── UI PRIMITIVES ────────────────────────────────────────────────────────────
const AV_C=["#2563EB","#7C3AED","#DB2777","#059669","#D97706","#DC2626","#0891B2","#EA580C"];
function avC(s){let h=0;for(let c of(s||""))h=(h*31+c.charCodeAt(0))&0xffff;return AV_C[h%AV_C.length];}
function initials(n){return(n||"?").split(/[\s,]+/).filter(Boolean).map(w=>w[0]).join("").slice(0,2).toUpperCase();}
function Avatar({name,size=36}){return <div style={{width:size,height:size,borderRadius:"50%",background:avC(name),color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*0.36,fontWeight:700,flexShrink:0,userSelect:"none"}}>{initials(name)}</div>;}

const STATUS_STYLES={Active:{bg:C.successL,c:C.successD},"On Leave":{bg:C.warningL,c:C.warningD},Present:{bg:C.successL,c:C.successD},Absent:{bg:C.dangerL,c:C.dangerD},Pending:{bg:C.warningL,c:C.warningD},Approved:{bg:C.successL,c:C.successD},Rejected:{bg:C.dangerL,c:C.dangerD},"Pending Supervisor":{bg:"#FEF3C7",c:"#92400E"},"Pending HR":{bg:C.purpleL,c:C.purple},Published:{bg:C.successL,c:C.successD},Draft:{bg:"#F1F5F9",c:"#475569"},Processed:{bg:C.tealL,c:C.teal},Upcoming:{bg:C.accentL,c:C.accentT},Completed:{bg:C.successL,c:C.successD},Cancelled:{bg:C.dangerL,c:C.dangerD},Open:{bg:C.warningL,c:C.warningD},Closed:{bg:"#F1F5F9",c:"#475569"},Inactive:{bg:"#F1F5F9",c:"#475569"},SC:{bg:"#DBEAFE",c:"#1E40AF"},SPR:{bg:"#E0E7FF",c:"#3730A3"},EP:{bg:C.warningL,c:C.warningD},WP:{bg:"#FCE7F3",c:"#9D174D"},superadmin:{bg:C.dangerL,c:C.dangerD},admin:{bg:"#DBEAFE",c:"#1E40AF"},supervisor:{bg:C.purpleL,c:C.purple},employee:{bg:"#F1F5F9",c:"#475569"},Safety:{bg:C.dangerL,c:C.dangerD},Holiday:{bg:C.warningL,c:C.warningD},Policy:{bg:C.purpleL,c:C.purple},General:{bg:C.accentL,c:C.accentT},Compliance:{bg:C.orangeL,c:C.orange},Urgent:{bg:C.dangerL,c:C.dangerD}};
function Badge({s}){const st=STATUS_STYLES[s]||{bg:"#F1F5F9",c:"#475569"};return <span style={{background:st.bg,color:st.c,borderRadius:99,padding:"2px 10px",fontSize:12,fontWeight:600,whiteSpace:"nowrap"}}>{s||"-"}</span>;}

function Card({children,style={},onClick}){return <div onClick={onClick} style={{background:C.card,borderRadius:12,border:`1px solid ${C.border}`,padding:"18px 22px",...(onClick?{cursor:"pointer"}:{}),...style}}>{children}</div>;}
function PageTitle({title,sub,actions}){return <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10,marginBottom:16}}><div><h1 style={{fontSize:19,fontWeight:800,margin:0,color:C.text}}>{title}</h1>{sub&&<p style={{color:C.muted,fontSize:13,margin:"2px 0 0"}}>{sub}</p>}</div>{actions&&<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{actions}</div>}</div>;}
function SecTitle({children}){return <div style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:10}}>{children}</div>;}

function StatCard({label,value,sub,icon,color}){return <Card style={{display:"flex",alignItems:"center",gap:12,flex:1,minWidth:130,padding:"14px 18px"}}><div style={{width:44,height:44,borderRadius:10,background:color+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:800,color,flexShrink:0}}>{icon}</div><div><div style={{fontSize:22,fontWeight:800,color:C.text,lineHeight:1}}>{value}</div><div style={{fontSize:12,color:C.muted,marginTop:2}}>{label}</div>{sub&&<div style={{fontSize:11,color,marginTop:1,fontWeight:600}}>{sub}</div>}</div></Card>;}

function Btn({children,onClick,v="primary",sm=false,style={},disabled=false}){
  const base={border:"none",borderRadius:7,fontWeight:600,cursor:disabled?"not-allowed":"pointer",opacity:disabled?0.6:1,fontSize:sm?12:13,padding:sm?"4px 11px":"8px 16px",fontFamily:"inherit",whiteSpace:"nowrap",...style};
  const V={primary:{background:C.accent,color:"#fff"},danger:{background:C.danger,color:"#fff"},ghost:{background:"transparent",color:C.muted,border:`1px solid ${C.border}`},success:{background:C.success,color:"#fff"},outline:{background:"transparent",color:C.accent,border:`1px solid ${C.accent}`},warning:{background:C.warning,color:"#fff"},teal:{background:C.teal,color:"#fff"},purple:{background:C.purple,color:"#fff"}};
  return <button disabled={disabled} onClick={e=>{e.stopPropagation();onClick&&onClick(e);}} style={{...base,...V[v]||V.primary}}>{children}</button>;
}

const IS={border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 11px",fontSize:13,outline:"none",background:"#fff",color:C.text,width:"100%",boxSizing:"border-box",fontFamily:"inherit"};
function Inp({label,value,onChange,type="text",placeholder="",required=false,error="",readOnly=false,rows}){
  const ctrl=rows
    ?<textarea value={value||""} onChange={e=>onChange&&onChange(e.target.value)} placeholder={placeholder} rows={rows} style={{...IS,resize:"vertical"}}/>
    :<input type={type} value={value||""} onChange={e=>onChange&&onChange(e.target.value)} placeholder={placeholder} readOnly={readOnly} style={{...IS,border:`1px solid ${error?C.danger:C.border}`,background:readOnly?"#F8FAFC":"#fff"}}/>;
  return <div style={{display:"flex",flexDirection:"column",gap:3}}>{label&&<label style={{fontSize:12,fontWeight:600,color:C.muted}}>{label}{required&&<span style={{color:C.danger}}> *</span>}</label>}{ctrl}{error&&<span style={{fontSize:11,color:C.danger}}>{error}</span>}</div>;
}
function Sel({label,value,onChange,options,required=false}){return <div style={{display:"flex",flexDirection:"column",gap:3}}>{label&&<label style={{fontSize:12,fontWeight:600,color:C.muted}}>{label}{required&&<span style={{color:C.danger}}> *</span>}</label>}<select value={value||""} onChange={e=>onChange&&onChange(e.target.value)} style={IS}><option value="">- Select -</option>{options.map(o=>typeof o==="string"?<option key={o} value={o}>{o}</option>:<option key={o.v} value={o.v}>{o.l}</option>)}</select></div>;}
function Grid({cols=2,children,style={}}){return <div style={{display:"grid",gridTemplateColumns:`repeat(${cols},1fr)`,gap:"12px 14px",...style}}>{children}</div>;}
function FormSection({title,children}){return <div style={{marginBottom:18}}><div style={{fontSize:11,fontWeight:700,color:C.accent,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:9,paddingBottom:5,borderBottom:`1px solid ${C.border}`}}>{title}</div>{children}</div>;}

function Modal({title,onClose,children,width=560}){return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={e=>e.target===e.currentTarget&&onClose()}><div style={{background:C.card,borderRadius:14,width:"100%",maxWidth:width,maxHeight:"93vh",overflowY:"auto",boxShadow:"0 24px 64px rgba(0,0,0,0.2)"}}><div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"16px 22px 12px",borderBottom:`1px solid ${C.border}`,position:"sticky",top:0,background:C.card,zIndex:1}}><h2 style={{margin:0,fontSize:15,fontWeight:800,color:C.text}}>{title}</h2><button onClick={onClose} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:C.muted,lineHeight:1,padding:0}}>x</button></div><div style={{padding:"18px 22px"}}>{children}</div></div></div>;}

function Toast({msg,type,onDone}){
  useEffect(()=>{const tm=setTimeout(onDone,3200);return()=>clearTimeout(tm);},[]);
  return <div style={{position:"fixed",bottom:24,right:24,zIndex:2000,background:type==="error"?C.danger:C.success,color:"#fff",borderRadius:9,padding:"11px 20px",fontWeight:600,fontSize:13,boxShadow:"0 8px 24px rgba(0,0,0,0.18)",display:"flex",gap:8,alignItems:"center"}}><span>{type==="error"?"✕":"✓"}</span><span>{msg}</span></div>;
}

function Confirm({msg,onOk,onCancel}){return <Modal title="Confirm" onClose={onCancel} width={360}><p style={{margin:"0 0 18px",color:C.text,fontSize:14,lineHeight:1.6}}>{msg}</p><div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={onCancel}>Cancel</Btn><Btn v="danger" onClick={onOk}>Confirm</Btn></div></Modal>;}

function TH({cols}){return <tr style={{background:C.bg}}>{cols.map(c=><th key={c} style={{padding:"10px 13px",textAlign:"left",fontWeight:700,color:C.muted,fontSize:11,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:"0.04em"}}>{c}</th>)}</tr>;}
function TR({children,onClick}){const[h,sH]=useState(false);return <tr style={{background:h?C.bg:"#fff",cursor:onClick?"pointer":"default"}} onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)} onClick={onClick}>{children}</tr>;}
function TD({children,style={}}){return <td style={{padding:"10px 13px",borderBottom:`1px solid ${C.border}`,...style}}>{children}</td>;}
function TTable({cols,rows,empty="No records found."}){return <Card style={{padding:0,overflow:"hidden"}}><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}><thead><TH cols={cols}/></thead><tbody>{rows.length===0?<tr><td colSpan={cols.length} style={{padding:32,textAlign:"center",color:C.muted,fontSize:13}}>{empty}</td></tr>:rows}</tbody></table></div></Card>;}

// Approval status chip with 3-level display
function ApprovalChain({status,lang}){
  const steps=[
    {label:t(lang,"pendingSup"),done:status!=="Pending Supervisor"},
    {label:t(lang,"pendingHR"),done:status==="Approved"||status==="Rejected"||status==="Pending HR"},
    {label:status==="Rejected"?t(lang,"Rejected"):t(lang,"Approved"),done:status==="Approved"||status==="Rejected"},
  ];
  return <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
    {steps.map((s,i)=><span key={i} style={{display:"flex",gap:4,alignItems:"center"}}>
      <span style={{fontSize:11,padding:"2px 8px",borderRadius:99,background:s.done?C.successL:C.bg,color:s.done?C.successD:C.muted,fontWeight:600}}>{s.label}</span>
      {i<2&&<span style={{fontSize:10,color:C.muted}}>→</span>}
    </span>)}
  </div>;
}

function MemoPopup({memos,empId,lang,onRead,onClose}){
  const unread=memos.filter(m=>!m.readBy.includes(empId));
  if(!unread.length)return null;
  const m=unread[0];
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <Card style={{maxWidth:480,width:"100%",border:`2px solid ${C.accent}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div style={{display:"flex",gap:8,alignItems:"center"}}><span style={{fontSize:20}}>📢</span><strong style={{fontSize:15}}>{lang==="zh"&&m.titleZh?m.titleZh:m.title}</strong></div>
        <Badge s={m.tag}/>
      </div>
      <p style={{fontSize:13,lineHeight:1.7,color:C.muted,margin:"0 0 16px"}}>{lang==="zh"&&m.bodyZh?m.bodyZh:m.body}</p>
      <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Posted by {m.author} · {fmtDate(m.date)}</div>
      {unread.length>1&&<div style={{fontSize:12,color:C.accent,marginBottom:10}}>+{unread.length-1} more announcement(s)</div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn v="outline" onClick={onClose}>Later</Btn>
        <Btn onClick={()=>onRead(m.id)}>Mark as Read</Btn>
      </div>
    </Card>
  </div>;
}

function TabBar({tabs,active,onChange}){
  return <div style={{display:"flex",gap:0,borderBottom:"1px solid "+C.border,marginBottom:16}}>
    {tabs.map(tab=>{
      const iA=active===tab.id;
      return <button key={tab.id} onClick={()=>onChange(tab.id)} style={{padding:"9px 16px",fontWeight:iA?700:500,fontSize:13,border:"none",background:"none",cursor:"pointer",color:iA?C.accent:C.muted,borderBottom:iA?"2px solid "+C.accent:"2px solid transparent",marginBottom:-1,fontFamily:"inherit",whiteSpace:"nowrap"}}>{tab.label}</button>;
    })}
  </div>;
}

function FilterBar({children}){return <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12}}>{children}</div>;}
function SearchInp({value,onChange,placeholder}){return <input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder||"Search..."} style={{flex:1,minWidth:160,...IS}}/>;}
function StatusPill({s,active,onClick}){const st=STATUS_STYLES[s]||{bg:C.bg,c:C.muted};return <button onClick={()=>onClick(s)} style={{border:`1px solid ${active?st.c:C.border}`,borderRadius:99,padding:"4px 14px",fontSize:12,fontWeight:active?700:500,background:active?st.bg:"#fff",color:active?st.c:C.muted,cursor:"pointer",fontFamily:"inherit"}}>{s}</button>;}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: ATTENDANCE (GPS + PHOTO)
// ═══════════════════════════════════════════════════════════════════════════════
function EmpAttendance({empId,attendance,onClock,lang}){
  const [gps,setGps]=useState(null);
  const [gpsErr,setGpsErr]=useState("");
  const [photo,setPhoto]=useState(null);
  const [getting,setGetting]=useState(false);
  const fileRef=useRef();
  const today=todayStr();
  const todayRec=attendance.find(a=>a.empId===empId&&a.date===today);

  function getLocation(){
    setGetting(true);setGpsErr("");
    if(!navigator.geolocation){setGpsErr("GPS not supported on this device.");setGetting(false);return;}
    navigator.geolocation.getCurrentPosition(
      pos=>{setGps({lat:pos.coords.latitude.toFixed(5),lng:pos.coords.longitude.toFixed(5)});setGetting(false);},
      err=>{setGpsErr("Unable to get location. Please enable GPS.");setGetting(false);},
      {enableHighAccuracy:true,timeout:10000}
    );
  }

  function handlePhoto(e){
    const file=e.target.files?.[0];
    if(!file)return;
    const reader=new FileReader();
    reader.onload=ev=>setPhoto(ev.target.result);
    reader.readAsDataURL(file);
  }

  function clockIn(){
    if(!gps){alert("Please capture GPS location first.");return;}
    onClock({id:uid(),empId,date:today,clockIn:new Date().toTimeString().slice(0,5),clockOut:"",lat:parseFloat(gps.lat),lng:parseFloat(gps.lng),site:"",photo,status:"Present",hoursWorked:0});
    setGps(null);setPhoto(null);
  }
  function clockOut(){
    const ci=todayRec.clockIn.split(":").map(Number);
    const co=new Date().toTimeString().slice(0,5).split(":").map(Number);
    const hrs=((co[0]*60+co[1])-(ci[0]*60+ci[1]))/60;
    onClock({...todayRec,clockOut:new Date().toTimeString().slice(0,5),hoursWorked:Math.round(hrs*100)/100,lat:gps?parseFloat(gps.lat):todayRec.lat,lng:gps?parseFloat(gps.lng):todayRec.lng});
    setGps(null);setPhoto(null);
  }

  const myHistory=attendance.filter(a=>a.empId===empId).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,14);
  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    <PageTitle title={t(lang,"attendance")} sub={fmtDate(today)}/>
    {/* Clock panel */}
    <Card style={{borderTop:`4px solid ${C.accent}`}}>
      <SecTitle>{todayRec?.clockIn?"Today's Record":"Clock In / Out"}</SecTitle>
      {todayRec&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
        <div style={{background:C.successL,borderRadius:8,padding:"10px 14px"}}>
          <div style={{fontSize:11,color:C.successD}}>Clock In</div>
          <div style={{fontSize:22,fontWeight:800,color:C.successD}}>{todayRec.clockIn}</div>
        </div>
        <div style={{background:todayRec.clockOut?C.dangerL:C.bg,borderRadius:8,padding:"10px 14px"}}>
          <div style={{fontSize:11,color:C.dangerD}}>Clock Out</div>
          <div style={{fontSize:22,fontWeight:800,color:C.dangerD}}>{todayRec.clockOut||"—"}</div>
        </div>
      </div>}
      {(!todayRec||!todayRec.clockOut)&&<>
        <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap"}}>
          <Btn v="outline" onClick={getLocation} disabled={getting}>{getting?t(lang,"capturing"):t(lang,"gps")}</Btn>
          <Btn v="ghost" onClick={()=>fileRef.current?.click()}>{t(lang,"photo")} 📷</Btn>
          <input ref={fileRef} type="file" accept="image/*" capture="camera" style={{display:"none"}} onChange={handlePhoto}/>
        </div>
        {gpsErr&&<div style={{color:C.danger,fontSize:12,marginBottom:8}}>{gpsErr}</div>}
        {gps&&<div style={{background:C.tealL,borderRadius:8,padding:"8px 12px",fontSize:12,color:C.teal,marginBottom:10,fontWeight:600}}>
          ✓ {t(lang,"locationCaptured")}: {gps.lat}, {gps.lng}
        </div>}
        {photo&&<div style={{marginBottom:10}}><img src={photo} alt="clock photo" style={{width:120,height:80,objectFit:"cover",borderRadius:8,border:`1px solid ${C.border}`}}/></div>}
        <div style={{display:"flex",gap:8}}>
          {!todayRec&&<Btn onClick={clockIn}>{t(lang,"clockIn")}</Btn>}
          {todayRec&&!todayRec.clockOut&&<Btn v="danger" onClick={clockOut}>{t(lang,"clockOut")}</Btn>}
        </div>
      </>}
      {todayRec?.clockOut&&<div style={{color:C.success,fontWeight:600,fontSize:13}}>✓ Completed — {todayRec.hoursWorked}h worked today</div>}
    </Card>
    {/* History */}
    <TTable cols={["Date","Clock In","Clock Out","Hours","Site","Status"]}
      rows={myHistory.map((a,i)=><TR key={a.id}>
        <TD>{fmtDate(a.date)}</TD><TD style={{fontWeight:600}}>{a.clockIn||"—"}</TD>
        <TD>{a.clockOut||"—"}</TD><TD>{a.hoursWorked?a.hoursWorked+"h":"—"}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{a.site||"—"}</TD><TD><Badge s={a.status}/></TD>
      </TR>)}/>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: LEAVE
// ═══════════════════════════════════════════════════════════════════════════════
function EmpLeave({empId,leaves,leaveTypes,onAdd,lang}){
  const [tab,setTab]=useState("apply");
  const [form,setForm]=useState({type:leaveTypes[0]?.name||"Annual Leave",from:todayStr(),to:todayStr(),reason:""});
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const myLeaves=leaves.filter(l=>l.empId===empId).sort((a,b)=>b.submittedDate.localeCompare(a.submittedDate));
  const approved=myLeaves.filter(l=>l.status==="Approved");
  const alUsed=approved.filter(l=>l.type==="Annual Leave").reduce((s,l)=>s+l.days,0);
  const mlUsed=approved.filter(l=>l.type==="Medical Leave").reduce((s,l)=>s+l.days,0);
  const alEnt=leaveTypes.find(x=>x.name==="Annual Leave")?.days||14;
  const mlEnt=leaveTypes.find(x=>x.name==="Medical Leave")?.days||14;
  function calcDays(f,to){const ms=new Date(to)-new Date(f);return Math.max(1,Math.floor(ms/86400000)+1);}
  function submit(){
    if(!form.reason.trim())return;
    onAdd({...form,empId,days:calcDays(form.from,form.to),status:"Pending Supervisor",supervisorComment:"",hrComment:"",submittedDate:todayStr()});
    alert("Leave request submitted for supervisor approval.");
  }
  const TABS=[{id:"apply",label:t(lang,"applyLeave")},{id:"balance",label:t(lang,"leaveBalance")},{id:"history",label:"History"}];
  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    <PageTitle title={t(lang,"leaves")}/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="apply"&&<Card>
      <Grid><Sel label={t(lang,"type")} value={form.type} onChange={v=>set("type",v)} options={leaveTypes.map(l=>l.name)}/>
        <div/>
        <Inp label={t(lang,"from")} type="date" value={form.from} onChange={v=>set("from",v)}/>
        <Inp label={t(lang,"to")} type="date" value={form.to} onChange={v=>set("to",v)}/>
      </Grid>
      <div style={{background:C.accentL,borderRadius:7,padding:"8px 12px",fontSize:13,color:C.accentT,fontWeight:600,margin:"10px 0"}}>
        {t(lang,"days")}: {calcDays(form.from,form.to)}
      </div>
      <Inp label={t(lang,"reason")} value={form.reason} onChange={v=>set("reason",v)} rows={3}/>
      <div style={{fontSize:11,color:C.muted,marginTop:8}}>Approval: Employee → Supervisor → HR</div>
      <div style={{marginTop:12}}><Btn onClick={submit}>{t(lang,"submit")}</Btn></div>
    </Card>}
    {tab==="balance"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      {[{name:"Annual Leave",ent:alEnt,used:alUsed},{name:"Medical Leave",ent:mlEnt,used:mlUsed}].concat(leaveTypes.filter(l=>l.name!=="Annual Leave"&&l.name!=="Medical Leave").map(l=>({name:l.name,ent:l.days,used:approved.filter(x=>x.type===l.name).reduce((s,x)=>s+x.days,0)}))).map(lt=><Card key={lt.name} style={{padding:"14px 18px"}}>
        <div style={{fontWeight:700,fontSize:14,marginBottom:8}}>{lt.name}</div>
        <div style={{display:"flex",gap:16}}>
          <div><div style={{fontSize:11,color:C.muted}}>{t(lang,"entitlement")}</div><div style={{fontWeight:800,fontSize:18}}>{lt.ent>=999?"∞":lt.ent}</div></div>
          <div><div style={{fontSize:11,color:C.muted}}>{t(lang,"used")}</div><div style={{fontWeight:800,fontSize:18,color:C.danger}}>{lt.used}</div></div>
          <div><div style={{fontSize:11,color:C.muted}}>{t(lang,"balance")}</div><div style={{fontWeight:800,fontSize:18,color:C.success}}>{lt.ent>=999?"∞":lt.ent-lt.used}</div></div>
        </div>
      </Card>)}
    </div>}
    {tab==="history"&&<TTable cols={["Type","From","To","Days","Status","Submitted"]}
      rows={myLeaves.map((l,i)=><TR key={l.id}>
        <TD style={{fontWeight:600}}>{l.type}</TD><TD style={{fontSize:12}}>{fmtDate(l.from)}</TD>
        <TD style={{fontSize:12}}>{fmtDate(l.to)}</TD><TD style={{fontWeight:700}}>{l.days}</TD>
        <TD><Badge s={l.status}/></TD><TD style={{fontSize:12,color:C.muted}}>{fmtDate(l.submittedDate)}</TD>
      </TR>)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: CLAIMS
// ═══════════════════════════════════════════════════════════════════════════════
function EmpClaims({empId,claims,onAdd,lang}){
  const [tab,setTab]=useState("submit");
  const [form,setForm]=useState({type:"Transport",date:todayStr(),amount:"",description:"",receipt:""});
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const fileRef=useRef();
  const myClaims=claims.filter(c=>c.empId===empId).sort((a,b)=>b.submittedDate.localeCompare(a.submittedDate));
  function handleReceipt(e){const file=e.target.files?.[0];if(!file)return;const r=new FileReader();r.onload=ev=>set("receipt",ev.target.result);r.readAsDataURL(file);}
  function submit(){
    if(!form.amount||!form.description.trim())return;
    onAdd({...form,empId,amount:parseFloat(form.amount),status:"Pending Supervisor",supervisorComment:"",hrComment:"",submittedDate:todayStr()});
    setForm({type:"Transport",date:todayStr(),amount:"",description:"",receipt:""});
    alert("Claim submitted for approval.");
  }
  const CLAIM_TYPES=[t(lang,"ot"),t(lang,"transport"),t(lang,"medical"),t(lang,"meal"),t(lang,"others")];
  const TABS=[{id:"submit",label:t(lang,"submitClaim")},{id:"history",label:"History"}];
  const pendingAmt=myClaims.filter(c=>c.status!=="Rejected").reduce((s,c)=>s+(c.status==="Approved"?c.amount:0),0);
  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    <PageTitle title={t(lang,"claims")}
      sub={`Approved: ${sgd(pendingAmt)}`}/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="submit"&&<Card><Grid>
      <Sel label={t(lang,"claimType")} value={form.type} onChange={v=>set("type",v)} options={CLAIM_TYPES}/>
      <Inp label={t(lang,"date")} type="date" value={form.date} onChange={v=>set("date",v)}/>
      <Inp label={`${t(lang,"amount")} (S$)`} type="number" value={form.amount} onChange={v=>set("amount",v)} required/>
      <div/>
    </Grid>
    <div style={{margin:"10px 0"}}><Inp label={t(lang,"description")} value={form.description} onChange={v=>set("description",v)} rows={2}/></div>
    <div style={{marginBottom:12}}>
      <label style={{fontSize:12,fontWeight:600,color:C.muted,display:"block",marginBottom:4}}>{t(lang,"receipt")}</label>
      <Btn v="ghost" sm onClick={()=>fileRef.current?.click()}>📎 Attach</Btn>
      <input ref={fileRef} type="file" accept="image/*,application/pdf" style={{display:"none"}} onChange={handleReceipt}/>
      {form.receipt&&<div style={{fontSize:11,color:C.success,marginTop:4}}>✓ File attached</div>}
    </div>
    <div style={{fontSize:11,color:C.muted,marginBottom:8}}>Approval: Employee → Supervisor → HR</div>
    <Btn onClick={submit}>{t(lang,"submit")}</Btn></Card>}
    {tab==="history"&&<TTable cols={["Type","Date","Amount","Description","Status"]}
      rows={myClaims.map(c=><TR key={c.id}>
        <TD><Badge s={c.type}/></TD><TD style={{fontSize:12}}>{fmtDate(c.date)}</TD>
        <TD style={{fontWeight:700,color:C.success}}>{sgd(c.amount)}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{c.description.slice(0,40)}</TD>
        <TD><Badge s={c.status}/></TD>
      </TR>)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: PAYSLIP
// ═══════════════════════════════════════════════════════════════════════════════
function EmpPayslip({empId,payroll,employees,lang}){
  const emp=employees.find(e=>e.id===empId)||{};
  const [selected,setSelected]=useState(null);
  const myPayslips=payroll.filter(p=>p.empId===empId&&p.status==="Published").sort((a,b)=>b.month.localeCompare(a.month));
  const p=selected||myPayslips[0];
  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    <PageTitle title={t(lang,"payslip")}/>
    {myPayslips.length===0?<Card><p style={{color:C.muted,textAlign:"center",padding:24}}>No payslips published yet.</p></Card>:
    <div style={{display:"grid",gridTemplateColumns:"220px 1fr",gap:16}}>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {myPayslips.map(pp=><button key={pp.id} onClick={()=>setSelected(pp)} style={{background:p?.id===pp.id?C.accent:"#fff",color:p?.id===pp.id?"#fff":C.text,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",cursor:"pointer",textAlign:"left",fontFamily:"inherit"}}>
          <div style={{fontWeight:700,fontSize:14}}>{pp.month}</div>
          <div style={{fontSize:12,opacity:0.8}}>{sgd(pp.netPay)}</div>
        </button>)}
      </div>
      {p&&<Card>
        <div style={{background:C.sidebar,color:"#fff",borderRadius:10,padding:"16px 20px",marginBottom:16}}>
          <div style={{fontSize:12,opacity:0.6,marginBottom:2}}>Custera.HR · Payslip</div>
          <div style={{fontWeight:800,fontSize:17}}>{emp.name}</div>
          <div style={{fontSize:12,opacity:0.7}}>{emp.position} · {emp.department}</div>
          <div style={{fontSize:12,opacity:0.7,marginTop:4}}>Period: {p.month} · FIN: {emp.fin||"-"}</div>
        </div>
        {[["Basic Salary",p.basic,null,true],["Allowance",p.allowance,null,false],["OT Pay",p.otPay||0,null,false],["Gross Pay",p.gross,null,true],["CPF (Employee 20%)",p.cpfEmployee,C.danger,false],["Net Pay",p.netPay,C.success,true]].map(([l,v,c,bold])=><div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.border}`,fontWeight:bold?800:400}}><span style={{color:C.muted,fontSize:13}}>{l}</span><span style={{color:c||C.text,fontSize:bold?15:13}}>{sgd(v)}</span></div>)}
        <div style={{marginTop:12,background:C.bg,borderRadius:8,padding:"10px 14px"}}>
          <div style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",marginBottom:6}}>Employer Contributions</div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:13}}><span style={{color:C.muted}}>CPF (Employer 17%)</span><span style={{color:C.warning,fontWeight:600}}>{sgd(p.cpfEmployer)}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginTop:4}}><span style={{color:C.muted}}>Skills Development Levy</span><span style={{fontWeight:600}}>{sgd(p.sdl)}</span></div>
        </div>
        <div style={{fontSize:11,color:C.muted,marginTop:12,textAlign:"center"}}>This is a computer-generated payslip. For queries, contact HR.</div>
      </Card>}
    </div>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: HR MEMO
// ═══════════════════════════════════════════════════════════════════════════════
function EmpMemo({memos,empId,lang,onRead}){
  const [sel,setSel]=useState(null);
  const sorted=[...memos].sort((a,b)=>b.date.localeCompare(a.date));
  const tagC={Safety:C.danger,Holiday:C.warning,Policy:C.purple,General:C.accent,Compliance:C.orange,Urgent:C.danger};
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"memo")}/>
    {sorted.map(m=>{const unread=!m.readBy.includes(empId);return <Card key={m.id} style={{borderLeft:`4px solid ${tagC[m.tag]||C.accent}`,cursor:"pointer",opacity:unread?1:0.75}} onClick={()=>{setSel(m);if(unread)onRead(m.id);}}>
      <div style={{display:"flex",justifyContent:"space-between",gap:12,alignItems:"flex-start"}}>
        <div>
          {unread&&<div style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:C.danger,marginRight:8,marginBottom:1}}/>}
          <strong style={{fontSize:14}}>{lang==="zh"&&m.titleZh?m.titleZh:m.title}</strong>
          <div style={{fontSize:12,color:C.muted,marginTop:4}}>{m.author} · {fmtDate(m.date)}</div>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center",flexShrink:0}}><Badge s={m.tag}/>{unread&&<span style={{background:C.dangerL,color:C.dangerD,fontSize:10,fontWeight:700,borderRadius:99,padding:"1px 6px"}}>NEW</span>}</div>
      </div>
    </Card>;})}
    {sel&&<Modal title={lang==="zh"&&sel.titleZh?sel.titleZh:sel.title} onClose={()=>setSel(null)} width={560}>
      <div style={{display:"flex",gap:8,marginBottom:12}}><Badge s={sel.tag}/><span style={{fontSize:12,color:C.muted}}>by {sel.author} · {fmtDate(sel.date)}</span></div>
      <p style={{lineHeight:1.8,fontSize:13,color:C.text,whiteSpace:"pre-wrap"}}>{lang==="zh"&&sel.bodyZh?sel.bodyZh:sel.body}</p>
    </Modal>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: COMPANY POLICY
// ═══════════════════════════════════════════════════════════════════════════════
function EmpPolicy({policies,lang}){
  const [sel,setSel]=useState(null);
  const cats=[...new Set(policies.map(p=>p.category))];
  const catC={Safety:C.danger,HR:C.purple,Finance:C.success,Operations:C.warning};
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"policy")}/>
    {cats.map(cat=><div key={cat}>
      <div style={{fontSize:12,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8}}>{cat}</div>
      {policies.filter(p=>p.category===cat).map(p=><Card key={p.id} style={{marginBottom:8,cursor:"pointer",borderLeft:`3px solid ${catC[cat]||C.accent}`}} onClick={()=>setSel(p)}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><strong style={{fontSize:14}}>{lang==="zh"&&p.titleZh?p.titleZh:p.title}</strong><div style={{fontSize:12,color:C.muted,marginTop:3}}>{p.version} · Effective {fmtDate(p.effectiveDate)}</div></div>
          <Btn v="ghost" sm onClick={()=>setSel(p)}>{t(lang,"view")}</Btn>
        </div>
      </Card>)}
    </div>)}
    {sel&&<Modal title={lang==="zh"&&sel.titleZh?sel.titleZh:sel.title} onClose={()=>setSel(null)} width={600}>
      <div style={{display:"flex",gap:8,marginBottom:12}}><Badge s={sel.category}/><span style={{fontSize:12,color:C.muted}}>{sel.version} · Effective {fmtDate(sel.effectiveDate)}</span></div>
      <pre style={{fontSize:13,lineHeight:1.8,whiteSpace:"pre-wrap",fontFamily:"inherit",color:C.text,background:C.bg,padding:"14px",borderRadius:8}}>{lang==="zh"&&sel.contentZh?sel.contentZh:sel.content}</pre>
    </Modal>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: TRAINING SCHEDULE
// ═══════════════════════════════════════════════════════════════════════════════
function EmpTraining({trainings,empId,lang}){
  const [sel,setSel]=useState(null);
  const myTrainings=trainings.filter(tr=>tr.assignedTo.includes(empId)||tr.assignedTo.length===0).sort((a,b)=>a.date.localeCompare(b.date));
  const statusC={Upcoming:C.accent,Completed:C.success,Cancelled:C.danger};
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"training")} sub={`${myTrainings.filter(t=>t.status==="Upcoming").length} upcoming`}/>
    {myTrainings.map(tr=><Card key={tr.id} style={{cursor:"pointer",borderLeft:`4px solid ${statusC[tr.status]||C.accent}`}} onClick={()=>setSel(tr)}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div>
          <strong style={{fontSize:14}}>{lang==="zh"&&tr.titleZh?tr.titleZh:tr.title}</strong>
          <div style={{display:"flex",gap:12,marginTop:6,fontSize:12,color:C.muted,flexWrap:"wrap"}}>
            <span>📅 {fmtDate(tr.date)} {tr.time}</span><span>⏱ {tr.duration}</span><span>📍 {tr.venue}</span>
          </div>
        </div>
        <Badge s={tr.status}/>
      </div>
    </Card>)}
    {myTrainings.length===0&&<Card><p style={{color:C.muted,textAlign:"center",padding:24}}>No training scheduled for you.</p></Card>}
    {sel&&<Modal title={lang==="zh"&&sel.titleZh?sel.titleZh:sel.title} onClose={()=>setSel(null)} width={520}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        {[["Date",fmtDate(sel.date)],["Time",sel.time],["Duration",sel.duration],["Venue",sel.venue],["Trainer",sel.trainer],["Status",sel.status]].map(([k,v])=><div key={k} style={{background:C.bg,borderRadius:7,padding:"8px 12px"}}><div style={{fontSize:11,color:C.muted}}>{k}</div><div style={{fontWeight:600,fontSize:13}}>{v}</div></div>)}
      </div>
      <p style={{fontSize:13,lineHeight:1.7,color:C.muted}}>{lang==="zh"&&sel.descriptionZh?sel.descriptionZh:sel.description}</p>
    </Modal>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: FEEDBACK
// ═══════════════════════════════════════════════════════════════════════════════
function EmpFeedback({feedbacks,sections,empId,onAdd,lang}){
  const [tab,setTab]=useState("submit");
  const [form,setForm]=useState({sectionId:"",subject:"",message:"",isAnonymous:false});
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const myFeedbacks=feedbacks.filter(f=>f.empId===empId).sort((a,b)=>b.date.localeCompare(a.date));
  function submit(){
    if(!form.sectionId||!form.subject.trim()||!form.message.trim())return;
    onAdd({...form,sectionId:parseInt(form.sectionId),empId:form.isAnonymous?"anon":empId,date:todayStr(),status:"Open",adminReply:""});
    setForm({sectionId:"",subject:"",message:"",isAnonymous:false});
    alert("Feedback submitted. Thank you!");
  }
  const TABS=[{id:"submit",label:t(lang,"submit")},{id:"history",label:"My Submissions"}];
  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    <PageTitle title={t(lang,"feedback")}/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="submit"&&<Card>
      <Sel label={t(lang,"category")} value={form.sectionId} onChange={v=>set("sectionId",v)} required
        options={sections.filter(s=>s.active).map(s=>({v:String(s.id),l:lang==="zh"&&s.nameZh?s.nameZh:s.name}))}/>
      <div style={{margin:"10px 0"}}><Inp label="Subject" value={form.subject} onChange={v=>set("subject",v)}/></div>
      <Inp label="Message" value={form.message} onChange={v=>set("message",v)} rows={4} placeholder="Share your feedback or idea..."/>
      <label style={{display:"flex",gap:8,alignItems:"center",fontSize:13,margin:"12px 0",cursor:"pointer"}}>
        <input type="checkbox" checked={form.isAnonymous} onChange={e=>set("isAnonymous",e.target.checked)}/>{t(lang,"anonymous")}
      </label>
      <Btn onClick={submit}>{t(lang,"submit")}</Btn>
    </Card>}
    {tab==="history"&&<TTable cols={["Category","Subject","Status","Admin Reply","Date"]}
      rows={myFeedbacks.map(f=>{const sec=sections.find(s=>s.id===f.sectionId);return <TR key={f.id}>
        <TD style={{fontSize:12}}>{lang==="zh"&&sec?.nameZh?sec.nameZh:sec?.name||"-"}</TD>
        <TD style={{fontWeight:600}}>{f.subject}</TD><TD><Badge s={f.status}/></TD>
        <TD style={{fontSize:12,color:C.muted}}>{f.adminReply||"—"}</TD><TD style={{fontSize:12,color:C.muted}}>{fmtDate(f.date)}</TD>
      </TR>;})}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: APPRAISAL
// ═══════════════════════════════════════════════════════════════════════════════
function EmpAppraisal({forms,submissions,empId,onSave,lang}){
  const [sel,setSel]=useState(null);
  const [draft,setDraft]=useState({});
  const mySubmissions=submissions.filter(s=>s.empId===empId);
  function getSubmission(formId){return mySubmissions.find(s=>s.formId===formId);}
  function saveDraft(form){
    const existing=getSubmission(form.id);
    const data={id:existing?.id||uid(),formId:form.id,empId,selfAnswers:draft,status:"Draft",submittedDate:"",supervisorAnswers:{},supervisorComment:"",finalStatus:""};
    onSave(data);setSel(null);
  }
  function submitForm(form){
    const existing=getSubmission(form.id);
    const data={id:existing?.id||uid(),formId:form.id,empId,selfAnswers:draft,status:"Submitted",submittedDate:todayStr(),supervisorAnswers:{},supervisorComment:"",finalStatus:""};
    onSave(data);setSel(null);alert("Appraisal submitted to your supervisor.");
  }
  const activeF=forms.filter(f=>f.status==="Active");
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"appraisal")}/>
    {activeF.map(f=>{const sub=getSubmission(f.id);return <Card key={f.id} style={{cursor:"pointer",borderLeft:`4px solid ${sub?C.success:C.accent}`}} onClick={()=>{setSel(f);setDraft(sub?.selfAnswers||{});}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div><strong style={{fontSize:14}}>{lang==="zh"&&f.titleZh?f.titleZh:f.title}</strong><div style={{fontSize:12,color:C.muted,marginTop:4}}>Period: {f.period} · Deadline: {fmtDate(f.deadline)}</div></div>
        <Badge s={sub?.status||"Not Started"}/>
      </div>
      {sub?.supervisorComment&&<div style={{marginTop:10,background:C.purpleL,borderRadius:7,padding:"8px 12px",fontSize:12,color:C.purple}}><strong>Supervisor feedback:</strong> {sub.supervisorComment}</div>}
    </Card>;})}
    {activeF.length===0&&<Card><p style={{color:C.muted,textAlign:"center",padding:24}}>No active appraisal forms.</p></Card>}
    {sel&&<Modal title={lang==="zh"&&sel.titleZh?sel.titleZh:sel.title} onClose={()=>setSel(null)} width={640}>
      <div style={{fontSize:12,color:C.muted,marginBottom:16}}>Period: {sel.period} · Deadline: {fmtDate(sel.deadline)}</div>
      {sel.questions.map(q=><div key={q.id} style={{marginBottom:16}}>
        <div style={{fontWeight:600,fontSize:13,marginBottom:6}}>{lang==="zh"&&q.questionZh?q.questionZh:q.question}</div>
        {q.type==="rating"?<div style={{display:"flex",gap:6}}>{[1,2,3,4,5].map(n=><button key={n} onClick={()=>setDraft(p=>({...p,[q.id]:n}))} style={{width:38,height:38,borderRadius:8,border:`2px solid ${draft[q.id]===n?C.accent:C.border}`,background:draft[q.id]===n?C.accentL:"#fff",fontWeight:700,fontSize:15,cursor:"pointer",color:draft[q.id]===n?C.accent:C.muted}}>{n}</button>)}<span style={{fontSize:11,color:C.muted,alignSelf:"center",marginLeft:4}}>1=Poor, 5=Excellent</span></div>
        :<textarea value={draft[q.id]||""} onChange={e=>setDraft(p=>({...p,[q.id]:e.target.value}))} rows={3} placeholder="Your response..." style={{...IS,resize:"vertical"}}/>}
      </div>)}
      <div style={{display:"flex",gap:8,marginTop:8}}>
        <Btn v="ghost" onClick={()=>saveDraft(sel)}>Save Draft</Btn>
        <Btn onClick={()=>submitForm(sel)}>{t(lang,"submit")}</Btn>
      </div>
    </Modal>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: COMPANY CALENDAR
// ═══════════════════════════════════════════════════════════════════════════════
function EmpCalendar({events,leaves,employees,empId,employees:emps,lang}){
  const [month,setMonth]=useState(new Date().toISOString().slice(0,7));
  const y=parseInt(month.split("-")[0]),m=parseInt(month.split("-")[1])-1;
  const daysInMonth=new Date(y,m+1,0).getDate();
  const firstDay=new Date(y,m,1).getDay();
  const cells=[];
  for(let i=0;i<firstDay;i++)cells.push(null);
  for(let i=1;i<=daysInMonth;i++)cells.push(i);
  function getEventsForDay(day){
    const ds=`${month}-${String(day).padStart(2,"0")}`;
    const evts=events.filter(e=>e.date<=ds&&e.endDate>=ds);
    const lvs=leaves.filter(l=>l.status==="Approved"&&l.from<=ds&&l.to>=ds);
    return{evts,lvs};
  }
  const DAYS=["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"calendar")} actions={[
      <Btn v="ghost" sm onClick={()=>{const d=new Date(y,m-1,1);setMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}}>&#8592;</Btn>,
      <span style={{fontSize:14,fontWeight:700,padding:"0 8px"}}>{new Date(y,m).toLocaleDateString("en-SG",{month:"long",year:"numeric"})}</span>,
      <Btn v="ghost" sm onClick={()=>{const d=new Date(y,m+1,1);setMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}}>&#8594;</Btn>,
    ]}/>
    <Card style={{padding:"14px"}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>
        {DAYS.map(d=><div key={d} style={{textAlign:"center",fontSize:11,fontWeight:700,color:C.muted,padding:"4px 0"}}>{d}</div>)}
        {cells.map((day,i)=>{
          if(!day)return <div key={`e${i}`}/>;
          const{evts,lvs}=getEventsForDay(day);
          const ds=`${month}-${String(day).padStart(2,"0")}`;
          const isToday=ds===todayStr();
          return <div key={day} style={{minHeight:72,background:isToday?C.accentL:"#F8FAFC",borderRadius:8,padding:"4px",border:`1px solid ${isToday?C.accent:C.border}`}}>
            <div style={{fontSize:12,fontWeight:isToday?800:600,color:isToday?C.accent:C.text,marginBottom:3}}>{day}</div>
            {evts.map(e=><div key={e.id} style={{fontSize:10,background:e.color||C.accent,color:"#fff",borderRadius:3,padding:"1px 4px",marginBottom:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.title}</div>)}
            {lvs.slice(0,2).map(l=>{const emp=emps.find(e=>e.id===l.empId);return <div key={l.id} style={{fontSize:10,background:C.warningL,color:C.warningD,borderRadius:3,padding:"1px 4px",marginBottom:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{emp?.name?.split(",")[0]||l.empId}</div>;})}
          </div>;
        })}
      </div>
    </Card>
    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
      {[{c:C.warning,l:"On Leave"},{c:C.accent,l:"Event"},{c:C.danger,l:"Holiday"}].map(x=><div key={x.l} style={{display:"flex",gap:6,alignItems:"center",fontSize:12}}><div style={{width:10,height:10,borderRadius:2,background:x.c}}/>{x.l}</div>)}
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE: SHIFT SCHEDULE
// ═══════════════════════════════════════════════════════════════════════════════
function EmpShift({shifts,empId,lang}){
  const myShifts=shifts.filter(s=>s.empId===empId).sort((a,b)=>b.weekStart.localeCompare(a.weekStart));
  const latest=myShifts[0];
  const shiftColor={Day:C.success,Night:C.purple,Off:C.muted,AL:C.warning};
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"shift")}/>
    {!latest?<Card><p style={{color:C.muted,textAlign:"center",padding:24}}>No shift schedule assigned yet.</p></Card>:
    <Card>
      <div style={{fontWeight:700,fontSize:14,marginBottom:14}}>Week of {fmtDate(latest.weekStart)}</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:6}}>
        {latest.shifts.map(s=>{const d=new Date(s.date);const day=d.toLocaleDateString("en-SG",{weekday:"short"});const dom=d.getDate();return <div key={s.date} style={{background:shiftColor[s.type]+"22",borderRadius:8,padding:"10px 6px",textAlign:"center",border:`1px solid ${shiftColor[s.type]||C.border}44`}}>
          <div style={{fontSize:11,fontWeight:700,color:C.muted}}>{day}</div>
          <div style={{fontSize:16,fontWeight:800,margin:"4px 0"}}>{dom}</div>
          <div style={{fontSize:11,fontWeight:700,color:shiftColor[s.type]||C.muted}}>{s.type}</div>
          {s.start&&<div style={{fontSize:10,color:C.muted,marginTop:2}}>{s.start}-{s.end}</div>}
          {s.site&&<div style={{fontSize:9,color:C.muted,marginTop:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.site}</div>}
        </div>;})}
      </div>
    </Card>}
    {myShifts.length>1&&<TTable cols={["Week","Shifts","Sites"]}
      rows={myShifts.slice(1).map(s=><TR key={s.id}>
        <TD>{fmtDate(s.weekStart)}</TD>
        <TD style={{fontSize:12}}>{s.shifts.map(sh=>`${new Date(sh.date).toLocaleDateString("en-SG",{weekday:"short"})}:${sh.type}`).join(", ")}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{[...new Set(s.shifts.map(sh=>sh.site).filter(Boolean))].join(", ")||"—"}</TD>
      </TR>)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: USER MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════
const MAX_ADMINS=3;
function UserMgmt({users,employees,onAdd,onEdit,onDelete,onResetPw,lang}){
  const [showForm,setShowForm]=useState(false);
  const [editU,setEditU]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const [showCreds,setShowCreds]=useState(null);
  const [f,setF]=useState({username:"",password:"",name:"",empId:"",role:"employee",active:true});
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  const adminCount=users.filter(u=>u.role==="admin"||u.role==="superadmin").length;
  function openNew(){setEditU(null);setF({username:"",password:genPwd(),name:"",empId:"",role:"employee",active:true});setShowForm(true);}
  function openEdit(u){setEditU(u);setF({...u});setShowForm(true);}
  function save(){
    if(!f.username.trim()||!f.name.trim()||!f.password||f.password.length<6)return;
    if(editU){onEdit(editU.id,f);}else{const nu={...f,id:"u"+uid()};onAdd(nu);setShowCreds(nu);}
    setShowForm(false);
  }
  function resetPw(u){const pw=genPwd();onResetPw(u.id,pw);setShowCreds({...u,password:pw,_reset:true});}
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title="User Accounts" sub={`${adminCount} admin accounts (max ${MAX_ADMINS+1} incl. super admin)`} actions={[<Btn onClick={openNew}>+ Create Account</Btn>]}/>
    <TTable cols={["User","Username","Role","Linked Employee","Status","Actions"]}
      rows={users.map((u,i)=><TR key={u.id}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={u.name} size={28}/><span style={{fontWeight:600}}>{u.name}</span></div></TD>
        <TD style={{fontSize:12}}>{u.username}</TD><TD><Badge s={u.role}/></TD>
        <TD style={{fontSize:12,color:C.muted}}>{employees.find(e=>e.id===u.empId)?.name||"—"}</TD>
        <TD><Badge s={u.active!==false?"Active":"Inactive"}/></TD>
        <TD><div style={{display:"flex",gap:4}}>
          {u.role!=="superadmin"&&<><Btn v="outline" sm onClick={()=>openEdit(u)}>{t(lang,"edit")}</Btn><Btn v="warning" sm onClick={()=>resetPw(u)}>Reset PW</Btn><Btn v="danger" sm onClick={()=>setConfirmId(u.id)}>{t(lang,"delete")}</Btn></>}
          {u.role==="superadmin"&&<span style={{fontSize:11,color:C.muted}}>Permanent</span>}
        </div></TD>
      </TR>)}/>
    {showForm&&<Modal title={editU?"Edit Account":"Create Account"} onClose={()=>setShowForm(false)} width={480}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Sel label="Link to Employee" value={f.empId} onChange={v=>{setFk("empId",v);const emp=employees.find(e=>e.id===v);if(emp&&!editU)setFk("name",emp.name);}} options={employees.map(e=>({v:e.id,l:`${e.id} - ${e.name}`}))}/>
        <Inp label="Display Name" value={f.name} onChange={v=>setFk("name",v)} required/>
        <Inp label="Username / Email" value={f.username} onChange={v=>setFk("username",v)} required/>
        <div style={{display:"flex",gap:8,alignItems:"flex-end"}}>
          <div style={{flex:1}}><Inp label="Password" value={f.password} onChange={v=>setFk("password",v)} required/></div>
          <Btn v="ghost" sm onClick={()=>setFk("password",genPwd())} style={{marginBottom:1}}>Generate</Btn>
        </div>
        <Sel label="Role" value={f.role} onChange={v=>setFk("role",v)} options={[{v:"employee",l:"Employee"},{v:"supervisor",l:"Supervisor / Reporting Officer"},{v:"admin",l:"Admin (HR)"}]}/>
        <label style={{display:"flex",gap:8,fontSize:13,cursor:"pointer"}}><input type="checkbox" checked={f.active!==false} onChange={e=>setFk("active",e.target.checked)}/>Account Active</label>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={save}>{editU?t(lang,"save"):"Create"}</Btn></div>
      </div>
    </Modal>}
    {showCreds&&<Modal title={showCreds._reset?"Password Reset":"Account Created"} onClose={()=>setShowCreds(null)} width={400}>
      <div style={{background:C.successL,borderRadius:8,padding:"12px 16px",marginBottom:14,fontSize:13,color:C.successD,fontWeight:600}}>Share these credentials securely with the employee:</div>
      <div style={{background:C.bg,borderRadius:8,padding:"14px",marginBottom:14}}>
        <div style={{fontSize:11,color:C.muted}}>Username</div><div style={{fontWeight:800,fontSize:16,marginBottom:8}}>{showCreds.username}</div>
        <div style={{fontSize:11,color:C.muted}}>Password</div><div style={{fontWeight:800,fontSize:16,fontFamily:"monospace",letterSpacing:2}}>{showCreds.password}</div>
      </div>
      <Btn onClick={()=>setShowCreds(null)}>Done</Btn>
    </Modal>}
    {confirmId&&<Confirm msg="Delete this user account?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: ATTENDANCE REVIEW
// ═══════════════════════════════════════════════════════════════════════════════
function AdminAttendance({attendance,employees,lang}){
  const [dateF,setDateF]=useState(todayStr());
  const [search,setSearch]=useState("");
  const recs=attendance.filter(a=>(!dateF||a.date===dateF)&&(!search||employees.find(e=>e.id===a.empId)?.name.toLowerCase().includes(search.toLowerCase())));
  const present=recs.filter(r=>r.status==="Present").length;
  const absent=employees.length-present;
  function doExport(){exportXLS(recs.map(a=>{const e=employees.find(x=>x.id===a.empId)||{};return{Date:fmtDate(a.date),"Employee ID":a.empId,Name:e.name||"",Site:a.site||"","Clock In":a.clockIn,"Clock Out":a.clockOut,"Hours Worked":a.hoursWorked,GPS:`${a.lat||""},${a.lng||""}`,Status:a.status};}),"Attendance",`attendance_${dateF||"all"}.xlsx`);}
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"attendance")} actions={[<Btn v="ghost" onClick={doExport}>{t(lang,"export")}</Btn>]}/>
    <div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
      <StatCard label="Present" value={present} icon="✓" color={C.success}/>
      <StatCard label="Absent/Unknown" value={absent} icon="✕" color={C.danger}/>
      <StatCard label="Records" value={recs.length} icon="R" color={C.accent}/>
    </div>
    <FilterBar>
      <SearchInp value={search} onChange={setSearch} placeholder="Search employee..."/>
      <input type="date" value={dateF} onChange={e=>setDateF(e.target.value)} style={{...IS,width:160}}/>
      <Btn v="ghost" sm onClick={()=>setDateF("")}>All Dates</Btn>
    </FilterBar>
    <TTable cols={["Employee","Date","Site","Clock In","Clock Out","Hours","GPS","Status"]}
      rows={recs.map(a=>{const emp=employees.find(e=>e.id===a.empId);return <TR key={a.id}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={emp?.name||"?"} size={26}/><span style={{fontWeight:600,fontSize:12}}>{emp?.name||a.empId}</span></div></TD>
        <TD style={{fontSize:12}}>{fmtDate(a.date)}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{a.site||"—"}</TD>
        <TD style={{fontWeight:600}}>{a.clockIn||"—"}</TD><TD>{a.clockOut||"—"}</TD>
        <TD style={{fontWeight:600,color:C.teal}}>{a.hoursWorked?a.hoursWorked+"h":"—"}</TD>
        <TD style={{fontSize:11,color:C.muted}}>{a.lat?`${a.lat},${a.lng}`:"—"}</TD>
        <TD><Badge s={a.status}/></TD>
      </TR>;})}/>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: LEAVE MANAGEMENT (3-level)
// ═══════════════════════════════════════════════════════════════════════════════
function AdminLeave({leaves,employees,leaveTypes,users,onAction,onAdd,onDelete,onAddType,onEditType,onDeleteType,session,lang}){
  const [tab,setTab]=useState("pending");
  const [filter,setFilter]=useState("Pending Supervisor");
  const [confirmId,setConfirmId]=useState(null);
  const isSupervisor=session.role==="supervisor";
  const myTeamIds=isSupervisor?users.filter(u=>u.reportingOfficerId===session.id).map(u=>u.empId):null;
  const visible=isSupervisor?leaves.filter(l=>myTeamIds.includes(l.empId)):leaves;
  const filtered=filter==="All"?visible:visible.filter(l=>l.status===filter);
  function approve(l){
    if(isSupervisor&&l.status==="Pending Supervisor"){onAction(l.id,"Pending HR",session.name,"");}
    else if(!isSupervisor&&l.status==="Pending HR"){onAction(l.id,"Approved",session.name,"");}
  }
  function reject(l,comment){onAction(l.id,"Rejected",session.name,comment||"");}
  const STATUSES=["All","Pending Supervisor","Pending HR","Approved","Rejected"];
  const TABS=[{id:"pending",label:"Leave Records"},{id:"types",label:"Leave Types"}];
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"leaves")} actions={[<Btn onClick={()=>{}} v="outline">+ Apply</Btn>]}/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="pending"&&<>
      <FilterBar>{STATUSES.map(s=><StatusPill key={s} s={s} active={filter===s} onClick={setFilter}/>)}</FilterBar>
      <TTable cols={["Employee","Type","From","To","Days","Reason","Status","Submitted","Actions"]}
        rows={filtered.map(l=>{const emp=employees.find(e=>e.id===l.empId);return <TR key={l.id}>
          <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={emp?.name||"?"} size={24}/><span style={{fontWeight:600,fontSize:12}}>{emp?.name||l.empId}</span></div></TD>
          <TD style={{fontSize:12}}>{l.type}</TD>
          <TD style={{fontSize:12,whiteSpace:"nowrap"}}>{fmtDate(l.from)}</TD>
          <TD style={{fontSize:12,whiteSpace:"nowrap"}}>{fmtDate(l.to)}</TD>
          <TD style={{fontWeight:700}}>{l.days}</TD>
          <TD style={{fontSize:12,color:C.muted,maxWidth:120}}><span title={l.reason}>{l.reason.slice(0,25)}{l.reason.length>25?"...":""}</span></TD>
          <TD><Badge s={l.status}/></TD>
          <TD style={{fontSize:11,color:C.muted}}>{fmtDate(l.submittedDate)}</TD>
          <TD><div style={{display:"flex",gap:4}}>
            {((isSupervisor&&l.status==="Pending Supervisor")||(!isSupervisor&&l.status==="Pending HR"))&&<>
              <Btn v="success" sm onClick={()=>approve(l)}>{t(lang,"approve")}</Btn>
              <Btn v="danger" sm onClick={()=>reject(l,"")}>{t(lang,"reject")}</Btn>
            </>}
            <Btn v="ghost" sm onClick={()=>setConfirmId(l.id)}>{t(lang,"delete")}</Btn>
          </div></TD>
        </TR>;})}/>
    </>}
    {tab==="types"&&<LeaveTypeAdmin types={leaveTypes} onAdd={onAddType} onEdit={onEditType} onDelete={onDeleteType} lang={lang}/>}
    {confirmId&&<Confirm msg="Delete leave record?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}
function LeaveTypeAdmin({types,onAdd,onEdit,onDelete,lang}){
  const [showForm,setShowForm]=useState(false);
  const [editIdx,setEditIdx]=useState(null);
  const [f,setF]=useState({name:"",days:"",paid:true,carryOver:false,desc:""});
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  function save(){if(!f.name.trim())return;const d={...f,days:Number(f.days)||0};editIdx!==null?onEdit(editIdx,d):onAdd(d);setShowForm(false);setEditIdx(null);setF({name:"",days:"",paid:true,carryOver:false,desc:""});}
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><SecTitle>Leave Types (Singapore MOM)</SecTitle><Btn sm onClick={()=>{setEditIdx(null);setF({name:"",days:"",paid:true,carryOver:false,desc:""});setShowForm(true);}}>+ Add Type</Btn></div>
    <TTable cols={["Leave Type","Days","Paid","Carry Over","Description",""]}
      rows={types.map((ty,i)=><TR key={i}>
        <TD style={{fontWeight:600}}>{ty.name}</TD><TD>{ty.days>=999?"Unlimited":ty.days}</TD>
        <TD><Badge s={ty.paid?"Active":"Inactive"}/></TD><TD style={{fontSize:12}}>{ty.carryOver?"Yes":"No"}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{ty.desc}</TD>
        <TD><div style={{display:"flex",gap:4}}><Btn v="outline" sm onClick={()=>{setF({...ty});setEditIdx(i);setShowForm(true);}}>{t(lang,"edit")}</Btn><Btn v="danger" sm onClick={()=>onDelete(i)}>{t(lang,"delete")}</Btn></div></TD>
      </TR>)}/>
    {showForm&&<Modal title={editIdx!==null?"Edit Leave Type":"Add Leave Type"} onClose={()=>setShowForm(false)} width={440}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Inp label="Leave Type Name" value={f.name} onChange={v=>setFk("name",v)} required/>
        <Inp label="Days (999 = unlimited)" type="number" value={f.days} onChange={v=>setFk("days",v)}/>
        <div style={{display:"flex",gap:20}}><label style={{display:"flex",gap:8,fontSize:13}}><input type="checkbox" checked={!!f.paid} onChange={e=>setFk("paid",e.target.checked)}/>Paid</label><label style={{display:"flex",gap:8,fontSize:13}}><input type="checkbox" checked={!!f.carryOver} onChange={e=>setFk("carryOver",e.target.checked)}/>Carry Over</label></div>
        <Inp label="Description" value={f.desc} onChange={v=>setFk("desc",v)} rows={2}/>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={save}>{t(lang,"save")}</Btn></div>
      </div>
    </Modal>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: CLAIMS MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════
function AdminClaims({claims,employees,users,session,onAction,onDelete,lang}){
  const [filter,setFilter]=useState("Pending Supervisor");
  const [confirmId,setConfirmId]=useState(null);
  const isSup=session.role==="supervisor";
  const myTeamIds=isSup?users.filter(u=>u.reportingOfficerId===session.id).map(u=>u.empId):null;
  const visible=isSup?claims.filter(c=>myTeamIds.includes(c.empId)):claims;
  const filtered=filter==="All"?visible:visible.filter(c=>c.status===filter);
  function approve(c){
    if(isSup&&c.status==="Pending Supervisor")onAction(c.id,"Pending HR",session.name,"");
    else if(!isSup&&c.status==="Pending HR")onAction(c.id,"Approved",session.name,"");
  }
  const totalApproved=visible.filter(c=>c.status==="Approved").reduce((s,c)=>s+c.amount,0);
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"claims")} sub={`Total Approved: ${sgd(totalApproved)}`} actions={[<Btn v="ghost" sm onClick={()=>exportXLS(filtered.map(c=>{const e=employees.find(x=>x.id===c.empId)||{};return{Name:e.name||"",Type:c.type,Date:fmtDate(c.date),"Amount (S$)":c.amount,Description:c.description,Status:c.status,"Supervisor":c.supervisorComment,"HR":c.hrComment};}),"Claims",`claims_${todayStr()}.xlsx`)}>{t(lang,"export")}</Btn>]}/>
    <FilterBar>{["All","Pending Supervisor","Pending HR","Approved","Rejected"].map(s=><StatusPill key={s} s={s} active={filter===s} onClick={setFilter}/>)}</FilterBar>
    <TTable cols={["Employee","Type","Date","Amount","Description","Status","Actions"]}
      rows={filtered.map(c=>{const emp=employees.find(e=>e.id===c.empId);return <TR key={c.id}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={emp?.name||"?"} size={24}/><span style={{fontWeight:600,fontSize:12}}>{emp?.name||c.empId}</span></div></TD>
        <TD><Badge s={c.type}/></TD><TD style={{fontSize:12}}>{fmtDate(c.date)}</TD>
        <TD style={{fontWeight:700,color:C.success}}>{sgd(c.amount)}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{c.description.slice(0,30)}</TD>
        <TD><Badge s={c.status}/></TD>
        <TD><div style={{display:"flex",gap:4}}>
          {((isSup&&c.status==="Pending Supervisor")||(!isSup&&c.status==="Pending HR"))&&<><Btn v="success" sm onClick={()=>approve(c)}>{t(lang,"approve")}</Btn><Btn v="danger" sm onClick={()=>onAction(c.id,"Rejected",session.name,"")}>{t(lang,"reject")}</Btn></>}
          <Btn v="ghost" sm onClick={()=>setConfirmId(c.id)}>{t(lang,"delete")}</Btn>
        </div></TD>
      </TR>;})}/>
    {confirmId&&<Confirm msg="Delete claim?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: HR MEMO
// ═══════════════════════════════════════════════════════════════════════════════
function AdminMemo({memos,onAdd,onEdit,onDelete,session,lang}){
  const [showForm,setShowForm]=useState(false);
  const [editM,setEditM]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const [f,setF]=useState({title:"",titleZh:"",body:"",bodyZh:"",tag:"General"});
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  function save(){if(!f.title.trim()||!f.body.trim())return;const data={...f,date:todayStr(),author:session.name,readBy:editM?.readBy||[]};editM?onEdit(editM.id,data):onAdd(data);setShowForm(false);setEditM(null);}
  function startEdit(m){setEditM(m);setF({title:m.title,titleZh:m.titleZh||"",body:m.body,bodyZh:m.bodyZh||"",tag:m.tag});setShowForm(true);}
  const TAGS=["General","Safety","Policy","Holiday","Compliance","Urgent"];
  const tagC={Safety:C.danger,Holiday:C.warning,Policy:C.purple,Urgent:C.danger,General:C.accent,Compliance:C.orange};
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"memo")} actions={[<Btn onClick={()=>{setEditM(null);setF({title:"",titleZh:"",body:"",bodyZh:"",tag:"General"});setShowForm(true);}}>+ {t(lang,"newMemo")}</Btn>]}/>
    {[...memos].sort((a,b)=>b.date.localeCompare(a.date)).map(m=><Card key={m.id} style={{borderLeft:`4px solid ${tagC[m.tag]||C.accent}`}}>
      <div style={{display:"flex",justifyContent:"space-between",gap:12}}>
        <div style={{flex:1}}><div style={{display:"flex",gap:8,alignItems:"center",marginBottom:6}}><strong style={{fontSize:14}}>{m.title}</strong><Badge s={m.tag}/></div>
          {m.titleZh&&<div style={{fontSize:12,color:C.muted,marginBottom:4}}>{m.titleZh}</div>}
          <p style={{fontSize:13,color:C.muted,margin:"0 0 6px",lineHeight:1.6}}>{m.body.slice(0,100)}{m.body.length>100?"...":""}</p>
          <div style={{fontSize:11,color:C.muted}}>by {m.author} · {fmtDate(m.date)} · {m.readBy.length} read</div>
        </div>
        <div style={{display:"flex",gap:4,flexShrink:0}}><Btn v="outline" sm onClick={()=>startEdit(m)}>{t(lang,"edit")}</Btn><Btn v="danger" sm onClick={()=>setConfirmId(m.id)}>{t(lang,"delete")}</Btn></div>
      </div>
    </Card>)}
    {showForm&&<Modal title={editM?"Edit Announcement":"New Announcement"} onClose={()=>setShowForm(false)} width={600}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Sel label="Tag" value={f.tag} onChange={v=>setFk("tag",v)} options={TAGS}/>
        <Grid><Inp label="Title (English)" value={f.title} onChange={v=>setFk("title",v)} required/><Inp label="标题 (中文)" value={f.titleZh} onChange={v=>setFk("titleZh",v)}/></Grid>
        <Grid><Inp label="Body (English)" value={f.body} onChange={v=>setFk("body",v)} rows={4}/><Inp label="内容 (中文)" value={f.bodyZh} onChange={v=>setFk("bodyZh",v)} rows={4}/></Grid>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={save}>{t(lang,"save")}</Btn></div>
      </div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete this announcement?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: COMPANY POLICY
// ═══════════════════════════════════════════════════════════════════════════════
function AdminPolicy({policies,onAdd,onEdit,onDelete,session,lang}){
  const [showForm,setShowForm]=useState(false);
  const [editP,setEditP]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const [f,setF]=useState({title:"",titleZh:"",category:"Safety",version:"v1.0",effectiveDate:"",content:"",contentZh:""});
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  function save(){if(!f.title.trim()&&!f.content.trim())return;const data={...f,uploadedBy:session.name,date:todayStr()};editP?onEdit(editP.id,data):onAdd(data);setShowForm(false);setEditP(null);}
  function startEdit(p){setEditP(p);setF({title:p.title,titleZh:p.titleZh||"",category:p.category,version:p.version,effectiveDate:p.effectiveDate,content:p.content,contentZh:p.contentZh||""});setShowForm(true);}
  const CATS=["Safety","HR","Finance","Operations","Quality","Environment"];
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"policy")} actions={[<Btn onClick={()=>{setEditP(null);setF({title:"",titleZh:"",category:"Safety",version:"v1.0",effectiveDate:"",content:"",contentZh:""});setShowForm(true);}}>+ Add Policy</Btn>]}/>
    <TTable cols={["Title","Category","Version","Effective","Uploaded By","Actions"]}
      rows={policies.map(p=><TR key={p.id}>
        <TD style={{fontWeight:600}}>{p.title}</TD><TD><Badge s={p.category}/></TD>
        <TD style={{fontSize:12}}>{p.version}</TD><TD style={{fontSize:12}}>{fmtDate(p.effectiveDate)}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{p.uploadedBy}</TD>
        <TD><div style={{display:"flex",gap:4}}><Btn v="outline" sm onClick={()=>startEdit(p)}>{t(lang,"edit")}</Btn><Btn v="danger" sm onClick={()=>setConfirmId(p.id)}>{t(lang,"delete")}</Btn></div></TD>
      </TR>)}/>
    {showForm&&<Modal title={editP?"Edit Policy":"Add Policy"} onClose={()=>setShowForm(false)} width={700}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Grid cols={3}><Sel label="Category" value={f.category} onChange={v=>setFk("category",v)} options={CATS}/><Inp label="Version" value={f.version} onChange={v=>setFk("version",v)}/><Inp label="Effective Date" type="date" value={f.effectiveDate} onChange={v=>setFk("effectiveDate",v)}/></Grid>
        <Grid><Inp label="Title (English)" value={f.title} onChange={v=>setFk("title",v)} required/><Inp label="标题 (中文)" value={f.titleZh} onChange={v=>setFk("titleZh",v)}/></Grid>
        <Grid><Inp label="Content (English)" value={f.content} onChange={v=>setFk("content",v)} rows={8}/><Inp label="内容 (中文)" value={f.contentZh} onChange={v=>setFk("contentZh",v)} rows={8}/></Grid>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={save}>{t(lang,"save")}</Btn></div>
      </div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete this policy?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: TRAINING
// ═══════════════════════════════════════════════════════════════════════════════
function AdminTraining({trainings,employees,onAdd,onEdit,onDelete,session,lang}){
  const [showForm,setShowForm]=useState(false);
  const [editT,setEditT]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const [f,setF]=useState({title:"",titleZh:"",description:"",descriptionZh:"",venue:"",trainer:"",date:"",time:"08:00",duration:"",targetDepts:[],assignedTo:[],status:"Upcoming"});
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  function save(){if(!f.title.trim())return;const data={...f};editT?onEdit(editT.id,data):onAdd(data);setShowForm(false);setEditT(null);}
  function startEdit(tr){setEditT(tr);setF({...tr});setShowForm(true);}
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"training")} sub={`${trainings.filter(tr=>tr.status==="Upcoming").length} upcoming sessions`} actions={[<Btn onClick={()=>{setEditT(null);setF({title:"",titleZh:"",description:"",descriptionZh:"",venue:"",trainer:"",date:"",time:"08:00",duration:"",targetDepts:[],assignedTo:[],status:"Upcoming"});setShowForm(true);}}>+ Add Training</Btn>]}/>
    <TTable cols={["Training","Date","Venue","Trainer","Assigned","Status","Actions"]}
      rows={trainings.map(tr=><TR key={tr.id}>
        <TD style={{fontWeight:600}}>{tr.title}</TD><TD style={{fontSize:12,whiteSpace:"nowrap"}}>{fmtDate(tr.date)} {tr.time}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{tr.venue}</TD><TD style={{fontSize:12}}>{tr.trainer}</TD>
        <TD style={{fontSize:12}}>{tr.assignedTo.length===0?"All":tr.assignedTo.length+" staff"}</TD>
        <TD><Badge s={tr.status}/></TD>
        <TD><div style={{display:"flex",gap:4}}><Btn v="outline" sm onClick={()=>startEdit(tr)}>{t(lang,"edit")}</Btn><Btn v="danger" sm onClick={()=>setConfirmId(tr.id)}>{t(lang,"delete")}</Btn></div></TD>
      </TR>)}/>
    {showForm&&<Modal title={editT?"Edit Training":"Add Training"} onClose={()=>setShowForm(false)} width={680}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Grid><Inp label="Title (English)" value={f.title} onChange={v=>setFk("title",v)} required/><Inp label="标题 (中文)" value={f.titleZh} onChange={v=>setFk("titleZh",v)}/></Grid>
        <Grid cols={3}><Inp label="Date" type="date" value={f.date} onChange={v=>setFk("date",v)}/><Inp label="Time" type="time" value={f.time} onChange={v=>setFk("time",v)}/><Inp label="Duration" value={f.duration} onChange={v=>setFk("duration",v)} placeholder="e.g. 8 hours"/></Grid>
        <Grid><Inp label="Venue" value={f.venue} onChange={v=>setFk("venue",v)}/><Inp label="Trainer" value={f.trainer} onChange={v=>setFk("trainer",v)}/></Grid>
        <Grid><Inp label="Description (EN)" value={f.description} onChange={v=>setFk("description",v)} rows={3}/><Inp label="描述 (中文)" value={f.descriptionZh} onChange={v=>setFk("descriptionZh",v)} rows={3}/></Grid>
        <Sel label="Status" value={f.status} onChange={v=>setFk("status",v)} options={["Upcoming","Completed","Cancelled"]}/>
        <div><label style={{fontSize:12,fontWeight:600,color:C.muted,display:"block",marginBottom:6}}>Assign Employees (leave empty for all)</label>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>{employees.map(e=><label key={e.id} style={{display:"flex",gap:4,fontSize:12,cursor:"pointer",background:f.assignedTo.includes(e.id)?C.accentL:C.bg,padding:"3px 8px",borderRadius:99,border:`1px solid ${f.assignedTo.includes(e.id)?C.accent:C.border}`}}><input type="checkbox" checked={f.assignedTo.includes(e.id)} onChange={ev=>setFk("assignedTo",ev.target.checked?[...f.assignedTo,e.id]:f.assignedTo.filter(x=>x!==e.id))}/>{e.name.split(",")[0]}</label>)}</div>
        </div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={save}>{t(lang,"save")}</Btn></div>
      </div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete training?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: FEEDBACK MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════
function AdminFeedback({feedbacks,sections,employees,onReply,onDeleteFeedback,onAddSection,onEditSection,onDeleteSection,lang}){
  const [tab,setTab]=useState("submissions");
  const [replyId,setReplyId]=useState(null);
  const [replyText,setReplyText]=useState("");
  const [showSectionForm,setShowSectionForm]=useState(false);
  const [editSec,setEditSec]=useState(null);
  const [sf,setSf]=useState({name:"",nameZh:"",description:"",active:true});
  function saveSection(){if(!sf.name.trim())return;editSec?onEditSection(editSec.id,sf):onAddSection(sf);setShowSectionForm(false);setEditSec(null);}
  const TABS=[{id:"submissions",label:"Submissions"},{id:"sections",label:"Categories"}];
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"feedback")} sub={`${feedbacks.filter(f=>f.status==="Open").length} open`}/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="submissions"&&<TTable cols={["Category","Subject","From","Message","Status","Reply","Date",""]}
      rows={feedbacks.map(f=>{const sec=sections.find(s=>s.id===f.sectionId);const emp=employees.find(e=>e.id===f.empId);return <TR key={f.id}>
        <TD style={{fontSize:12}}>{lang==="zh"&&sec?.nameZh?sec.nameZh:sec?.name||"—"}</TD>
        <TD style={{fontWeight:600,fontSize:12}}>{f.subject}</TD>
        <TD style={{fontSize:12}}>{f.empId==="anon"?"Anonymous":emp?.name||f.empId}</TD>
        <TD style={{fontSize:12,color:C.muted,maxWidth:150}}>{f.message.slice(0,50)}...</TD>
        <TD><Badge s={f.status}/></TD>
        <TD style={{fontSize:12,color:C.muted}}>{f.adminReply||"—"}</TD>
        <TD style={{fontSize:12}}>{fmtDate(f.date)}</TD>
        <TD><Btn v="outline" sm onClick={()=>{setReplyId(f.id);setReplyText(f.adminReply||"");}}>Reply</Btn></TD>
      </TR>;})}/>}
    {tab==="sections"&&<div>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}><Btn onClick={()=>{setEditSec(null);setSf({name:"",nameZh:"",description:"",active:true});setShowSectionForm(true);}}>+ Add Category</Btn></div>
      <TTable cols={["Category (EN)","Category (ZH)","Description","Active",""]}
        rows={sections.map(s=><TR key={s.id}>
          <TD style={{fontWeight:600}}>{s.name}</TD><TD style={{fontSize:12}}>{s.nameZh||"—"}</TD>
          <TD style={{fontSize:12,color:C.muted}}>{s.description}</TD>
          <TD><Badge s={s.active?"Active":"Inactive"}/></TD>
          <TD><div style={{display:"flex",gap:4}}><Btn v="outline" sm onClick={()=>{setEditSec(s);setSf({name:s.name,nameZh:s.nameZh||"",description:s.description,active:s.active});setShowSectionForm(true);}}>Edit</Btn><Btn v="danger" sm onClick={()=>onDeleteSection(s.id)}>Del</Btn></div></TD>
        </TR>)}/>
      {showSectionForm&&<Modal title={editSec?"Edit Category":"Add Category"} onClose={()=>setShowSectionForm(false)} width={480}>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <Grid><Inp label="Name (English)" value={sf.name} onChange={v=>setSf(p=>({...p,name:v}))} required/><Inp label="名称 (中文)" value={sf.nameZh} onChange={v=>setSf(p=>({...p,nameZh:v}))}/></Grid>
          <Inp label="Description" value={sf.description} onChange={v=>setSf(p=>({...p,description:v}))}/>
          <label style={{display:"flex",gap:8,fontSize:13}}><input type="checkbox" checked={sf.active} onChange={e=>setSf(p=>({...p,active:e.target.checked}))}/>Active</label>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowSectionForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={saveSection}>{t(lang,"save")}</Btn></div>
        </div>
      </Modal>}
    </div>}
    {replyId&&<Modal title="Reply to Feedback" onClose={()=>setReplyId(null)} width={480}>
      <Inp label="Your Reply" value={replyText} onChange={setReplyText} rows={4}/>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:12}}>
        <Btn v="ghost" onClick={()=>setReplyId(null)}>{t(lang,"cancel")}</Btn>
        <Btn onClick={()=>{onReply(replyId,replyText);setReplyId(null);}}>Send Reply & Close</Btn>
      </div>
    </Modal>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: APPRAISAL MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════
function AdminAppraisal({forms,submissions,employees,onAddForm,onEditForm,onDeleteForm,onSaveSubmission,session,lang}){
  const [tab,setTab]=useState("forms");
  const [showForm,setShowForm]=useState(false);
  const [editF,setEditF]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const [viewSub,setViewSub]=useState(null);
  const [f,setF]=useState({title:"",titleZh:"",period:"",deadline:"",targetDepts:[],status:"Draft",questions:[]});
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  function addQuestion(){setF(p=>({...p,questions:[...p.questions,{id:"q"+uid(),question:"",questionZh:"",type:"rating",category:"Performance"}]}))}
  function updateQ(i,k,v){setF(p=>({...p,questions:p.questions.map((q,x)=>x===i?{...q,[k]:v}:q)}));}
  function removeQ(i){setF(p=>({...p,questions:p.questions.filter((_,x)=>x!==i)}));}
  function save(){if(!f.title.trim())return;editF?onEditForm(editF.id,f):onAddForm(f);setShowForm(false);setEditF(null);}
  function supervisorReview(sub,comment,finalRatings){onSaveSubmission({...sub,supervisorComment:comment,supervisorAnswers:finalRatings,status:"Completed"});}
  const TABS=[{id:"forms",label:"Appraisal Forms"},{id:"submissions",label:"Submissions"}];
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"appraisal")} actions={[<Btn onClick={()=>{setEditF(null);setF({title:"",titleZh:"",period:"",deadline:"",targetDepts:[],status:"Draft",questions:[]});setShowForm(true);}}>+ Create Form</Btn>]}/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="forms"&&<TTable cols={["Title","Period","Deadline","Status","Responses","Actions"]}
      rows={forms.map(f2=><TR key={f2.id}>
        <TD style={{fontWeight:600}}>{f2.title}</TD><TD style={{fontSize:12}}>{f2.period}</TD>
        <TD style={{fontSize:12}}>{fmtDate(f2.deadline)}</TD><TD><Badge s={f2.status}/></TD>
        <TD>{submissions.filter(s=>s.formId===f2.id&&s.status==="Submitted").length}</TD>
        <TD><div style={{display:"flex",gap:4}}>
          <Btn v="outline" sm onClick={()=>{setEditF(f2);setF({...f2});setShowForm(true);}}>Edit</Btn>
          <Btn v={f2.status==="Active"?"warning":"success"} sm onClick={()=>onEditForm(f2.id,{...f2,status:f2.status==="Active"?"Closed":"Active"})}>{f2.status==="Active"?"Close":"Activate"}</Btn>
          <Btn v="danger" sm onClick={()=>setConfirmId(f2.id)}>Del</Btn>
        </div></TD>
      </TR>)}/>}
    {tab==="submissions"&&<TTable cols={["Employee","Form","Status","Self Submitted","Supervisor Review",""]}
      rows={submissions.map(s=>{const emp=employees.find(e=>e.id===s.empId);const frm=forms.find(f2=>f2.id===s.formId);return <TR key={s.id}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={emp?.name||"?"} size={24}/><span style={{fontWeight:600,fontSize:12}}>{emp?.name||s.empId}</span></div></TD>
        <TD style={{fontSize:12}}>{frm?.title||"—"}</TD><TD><Badge s={s.status}/></TD>
        <TD style={{fontSize:12}}>{fmtDate(s.submittedDate)}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{s.supervisorComment||"—"}</TD>
        <TD><Btn v="outline" sm onClick={()=>setViewSub(s)}>Review</Btn></TD>
      </TR>;})}/>}
    {viewSub&&<AppraisalReviewModal sub={viewSub} form={forms.find(f2=>f2.id===viewSub.formId)} emp={employees.find(e=>e.id===viewSub.empId)} onSave={supervisorReview} onClose={()=>setViewSub(null)} lang={lang}/>}
    {showForm&&<Modal title={editF?"Edit Form":"Create Appraisal Form"} onClose={()=>setShowForm(false)} width={700}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Grid><Inp label="Title (English)" value={f.title} onChange={v=>setFk("title",v)} required/><Inp label="标题 (中文)" value={f.titleZh} onChange={v=>setFk("titleZh",v)}/></Grid>
        <Grid><Inp label="Period (e.g. Jan-Jun 2026)" value={f.period} onChange={v=>setFk("period",v)}/><Inp label="Deadline" type="date" value={f.deadline} onChange={v=>setFk("deadline",v)}/></Grid>
        <Sel label="Status" value={f.status} onChange={v=>setFk("status",v)} options={["Draft","Active","Closed"]}/>
        <div><div style={{fontSize:12,fontWeight:600,color:C.muted,marginBottom:8}}>Questions</div>
          {f.questions.map((q,i)=><div key={q.id} style={{background:C.bg,borderRadius:8,padding:"10px 12px",marginBottom:8}}>
            <Grid cols={3} style={{marginBottom:8}}><Sel label="Type" value={q.type} onChange={v=>updateQ(i,"type",v)} options={["rating","text"]}/><Sel label="Category" value={q.category} onChange={v=>updateQ(i,"category",v)} options={["Performance","Safety","Behaviour","Attendance","Development"]}/><div style={{display:"flex",alignItems:"flex-end"}}><Btn v="danger" sm onClick={()=>removeQ(i)}>Remove</Btn></div></Grid>
            <Grid><Inp label="Question (EN)" value={q.question} onChange={v=>updateQ(i,"question",v)}/><Inp label="问题 (中文)" value={q.questionZh||""} onChange={v=>updateQ(i,"questionZh",v)}/></Grid>
          </div>)}
          <Btn v="ghost" sm onClick={addQuestion}>+ Add Question</Btn>
        </div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={save}>{t(lang,"save")}</Btn></div>
      </div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete appraisal form?" onOk={()=>{onDeleteForm(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}
function AppraisalReviewModal({sub,form,emp,onSave,onClose,lang}){
  const [comment,setComment]=useState(sub.supervisorComment||"");
  const [supRatings,setSupRatings]=useState(sub.supervisorAnswers||{});
  if(!form)return null;
  return <Modal title={`Review: ${emp?.name}`} onClose={onClose} width={680}>
    <div style={{marginBottom:14,background:C.bg,borderRadius:8,padding:"10px 14px",fontSize:13,color:C.muted}}>Period: {form.period} · Submitted: {fmtDate(sub.submittedDate)}</div>
    {form.questions.map(q=><div key={q.id} style={{marginBottom:16,paddingBottom:14,borderBottom:`1px solid ${C.border}`}}>
      <div style={{fontWeight:600,fontSize:13,marginBottom:6}}>{lang==="zh"&&q.questionZh?q.questionZh:q.question}</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div><div style={{fontSize:11,color:C.muted,marginBottom:4}}>Employee's Response</div>
          {q.type==="rating"?<div style={{display:"flex",gap:4}}>{[1,2,3,4,5].map(n=><div key={n} style={{width:30,height:30,borderRadius:6,background:sub.selfAnswers?.[q.id]===n?C.accentL:C.bg,border:`1px solid ${sub.selfAnswers?.[q.id]===n?C.accent:C.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:13,color:sub.selfAnswers?.[q.id]===n?C.accent:C.muted}}>{n}</div>)}</div>
          :<div style={{fontSize:12,color:C.text,background:C.bg,borderRadius:6,padding:"8px"}}>{sub.selfAnswers?.[q.id]||"—"}</div>}
        </div>
        {q.type==="rating"&&<div><div style={{fontSize:11,color:C.muted,marginBottom:4}}>Supervisor Rating</div>
          <div style={{display:"flex",gap:4}}>{[1,2,3,4,5].map(n=><button key={n} onClick={()=>setSupRatings(p=>({...p,[q.id]:n}))} style={{width:30,height:30,borderRadius:6,background:supRatings[q.id]===n?C.purpleL:C.bg,border:`1px solid ${supRatings[q.id]===n?C.purple:C.border}`,fontWeight:700,fontSize:13,cursor:"pointer",color:supRatings[q.id]===n?C.purple:C.muted}}>{n}</button>)}</div>
        </div>}
      </div>
    </div>)}
    <Inp label="Supervisor's Overall Comment" value={comment} onChange={setComment} rows={3}/>
    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:12}}><Btn v="ghost" onClick={onClose}>{t(lang,"cancel")}</Btn><Btn onClick={()=>{onSave(sub,comment,supRatings);onClose();}}>Submit Review</Btn></div>
  </Modal>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: COMPANY CALENDAR
// ═══════════════════════════════════════════════════════════════════════════════
function AdminCalendar({events,leaves,employees,onAdd,onEdit,onDelete,session,lang}){
  const [month,setMonth]=useState(new Date().toISOString().slice(0,7));
  const [showForm,setShowForm]=useState(false);
  const [editE,setEditE]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const [f,setF]=useState({title:"",type:"Meeting",date:"",endDate:"",dept:"All",color:C.accent});
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  const y=parseInt(month.split("-")[0]),m=parseInt(month.split("-")[1])-1;
  const daysInMonth=new Date(y,m+1,0).getDate();
  const firstDay=new Date(y,m,1).getDay();
  const cells=[];for(let i=0;i<firstDay;i++)cells.push(null);for(let i=1;i<=daysInMonth;i++)cells.push(i);
  function getDay(day){const ds=`${month}-${String(day).padStart(2,"0")}`;const evts=events.filter(e=>e.date<=ds&&e.endDate>=ds);const lvs=leaves.filter(l=>l.status==="Approved"&&l.from<=ds&&l.to>=ds);return{evts,lvs};}
  function save(){if(!f.title.trim()||!f.date)return;const data={...f,endDate:f.endDate||f.date,createdBy:session.name};editE?onEdit(editE.id,data):onAdd(data);setShowForm(false);setEditE(null);}
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"calendar")} actions={[
      <Btn v="ghost" sm onClick={()=>{const d=new Date(y,m-1,1);setMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}}>&#8592;</Btn>,
      <span style={{fontSize:14,fontWeight:700,padding:"0 8px"}}>{new Date(y,m).toLocaleDateString("en-SG",{month:"long",year:"numeric"})}</span>,
      <Btn v="ghost" sm onClick={()=>{const d=new Date(y,m+1,1);setMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}}>&#8594;</Btn>,
      <Btn onClick={()=>{setEditE(null);setF({title:"",type:"Meeting",date:`${month}-01`,endDate:`${month}-01`,dept:"All",color:C.accent});setShowForm(true);}}>+ Add Event</Btn>,
    ]}/>
    <Card style={{padding:"14px"}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>
        {["Sun","Mon","Tue","Wed","Thu","Fri","Sat"].map(d=><div key={d} style={{textAlign:"center",fontSize:11,fontWeight:700,color:C.muted,padding:"4px 0"}}>{d}</div>)}
        {cells.map((day,i)=>{if(!day)return <div key={`e${i}`}/>;const{evts,lvs}=getDay(day);const ds=`${month}-${String(day).padStart(2,"0")}`;const isToday=ds===todayStr();return <div key={day} style={{minHeight:72,background:isToday?C.accentL:"#F8FAFC",borderRadius:8,padding:"4px",border:`1px solid ${isToday?C.accent:C.border}`}}>
          <div style={{fontSize:12,fontWeight:isToday?800:600,color:isToday?C.accent:C.text,marginBottom:2}}>{day}</div>
          {evts.map(e=><div key={e.id} onClick={()=>{setEditE(e);setF({title:e.title,type:e.type,date:e.date,endDate:e.endDate,dept:e.dept||"All",color:e.color||C.accent});setShowForm(true);}} style={{fontSize:9,background:e.color||C.accent,color:"#fff",borderRadius:3,padding:"1px 4px",marginBottom:1,cursor:"pointer",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.title}</div>)}
          {lvs.slice(0,2).map(l=>{const emp=employees.find(e=>e.id===l.empId);return <div key={l.id} style={{fontSize:9,background:C.warningL,color:C.warningD,borderRadius:3,padding:"1px 4px",marginBottom:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{emp?.name?.split(",")[0]||l.empId}</div>;})}
        </div>;})}
      </div>
    </Card>
    {showForm&&<Modal title={editE?"Edit Event":"Add Event"} onClose={()=>setShowForm(false)} width={480}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Inp label="Title" value={f.title} onChange={v=>setFk("title",v)} required/>
        <Sel label="Type" value={f.type} onChange={v=>setFk("type",v)} options={["Holiday","Training","Meeting","Inspection","Milestone","Other"]}/>
        <Grid><Inp label="Start Date" type="date" value={f.date} onChange={v=>setFk("date",v)}/><Inp label="End Date" type="date" value={f.endDate} onChange={v=>setFk("endDate",v)}/></Grid>
        <Sel label="Department" value={f.dept} onChange={v=>setFk("dept",v)} options={["All","Management","Engineering & Operations","Business Development","HR & Admin.","Admin","Commercial"]}/>
        <div><label style={{fontSize:12,fontWeight:600,color:C.muted,display:"block",marginBottom:4}}>Color</label>
          <div style={{display:"flex",gap:8}}>{[C.accent,C.danger,C.warning,C.success,C.purple,C.orange,C.teal].map(c=><div key={c} onClick={()=>setFk("color",c)} style={{width:24,height:24,borderRadius:"50%",background:c,cursor:"pointer",border:f.color===c?"2px solid #000":"2px solid transparent"}}/> )}</div>
        </div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
          {editE&&<Btn v="danger" sm onClick={()=>{setConfirmId(editE.id);setShowForm(false);}}>Delete</Btn>}
          <Btn v="ghost" onClick={()=>setShowForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={save}>{t(lang,"save")}</Btn>
        </div>
      </div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete this event?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: PAYROLL (MOM Singapore)
// ═══════════════════════════════════════════════════════════════════════════════
function AdminPayroll({employees,payroll,claims,onProcess,onDelete,onPublish,lang}){
  const [month,setMonth]=useState(new Date().toISOString().slice(0,7));
  const [showSlip,setShowSlip]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const processed=payroll.filter(p=>p.month===month);
  const processedIds=new Set(processed.map(p=>p.empId));
  const activeEmps=employees.filter(e=>e.status==="Active");
  function processOne(e){
    const basic=Number(e.basicSalary)||0;const allow=Number(e.allowance)||0;
    const approvedOT=claims.filter(c=>c.empId===e.id&&c.type==="OT"&&c.status==="Approved"&&c.date.startsWith(month)).reduce((s,c)=>s+c.amount,0);
    const gross=basic+allow+approvedOT;const cpf=calcCPF(basic,e.nationality);const sdl=calcSDF(basic);const netPay=gross-cpf.employee;
    onProcess({id:uid(),empId:e.id,month,basic,allowance:allow,otPay:approvedOT,gross,cpfEmployee:cpf.employee,cpfEmployer:cpf.employer,sdl,netPay,status:"Draft",processedOn:todayStr(),publishedOn:""});
  }
  function processAll(){activeEmps.filter(e=>!processedIds.has(e.id)).forEach(processOne);}
  function doExport(){exportMulti([{n:"Payroll Summary",data:processed.map(p=>{const e=employees.find(x=>x.id===p.empId)||{};return{Month:p.month,"Emp ID":p.empId,Name:e.name||"",Company:e.company||"",Nationality:e.nationality||"","Work Pass":e.workPass||"","FIN/NRIC":e.fin||"","Basic (S$)":p.basic,"Allow (S$)":p.allowance,"OT (S$)":p.otPay||0,"Gross (S$)":p.gross,"CPF Ee (S$)":p.cpfEmployee,"CPF Er (S$)":p.cpfEmployer,"SDL (S$)":p.sdl,"Net Pay (S$)":p.netPay,Status:p.status};})},{n:"CPF Submission",data:processed.filter(p=>p.cpfEmployee>0).map(p=>{const e=employees.find(x=>x.id===p.empId)||{};return{Month:p.month,Name:e.name||"","FIN/NRIC":e.fin||""," OW":p.basic," AW":p.allowance,"Ee CPF":p.cpfEmployee,"Er CPF":p.cpfEmployer,"Total":p.cpfEmployee+p.cpfEmployer};})}],`payroll_${month}.xlsx`);}
  const totGross=processed.reduce((s,p)=>s+p.gross,0);const totNet=processed.reduce((s,p)=>s+p.netPay,0);const totEmpCPF=processed.reduce((s,p)=>s+p.cpfEmployee,0);const totErCPF=processed.reduce((s,p)=>s+p.cpfEmployer,0);
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:10}}>
      <div><h1 style={{fontSize:19,fontWeight:800,margin:0}}>{t(lang,"payroll")}</h1><input type="month" value={month} onChange={e=>setMonth(e.target.value)} style={{...IS,width:160,marginTop:6}}/></div>
      <div style={{display:"flex",gap:8}}><Btn v="ghost" onClick={doExport}>{t(lang,"export")}</Btn><Btn v="warning" onClick={processAll} disabled={activeEmps.filter(e=>!processedIds.has(e.id)).length===0}>Process All</Btn></div>
    </div>
    <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
      <StatCard label="Total Gross" value={sgd(totGross)} icon="$" color={C.accent}/>
      <StatCard label="Total Net Pay" value={sgd(totNet)} icon="N" color={C.success}/>
      <StatCard label="CPF (Employee)" value={sgd(totEmpCPF)} icon="E" color={C.purple}/>
      <StatCard label="CPF (Employer)" value={sgd(totErCPF)} icon="R" color={C.warning}/>
    </div>
    <TTable cols={["Employee","Nationality","Basic","OT","Gross","CPF Ee","CPF Er","SDL","Net Pay","Status",""]}
      rows={activeEmps.map(e=>{const p=processed.find(x=>x.empId===e.id);const basic=Number(e.basicSalary)||0;const allow=Number(e.allowance)||0;const cpf=calcCPF(basic,e.nationality);const pendingOT=claims.filter(c=>c.empId===e.id&&c.type==="OT"&&c.status==="Approved"&&c.date.startsWith(month)).reduce((s,c)=>s+c.amount,0);
      return <TR key={e.id}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={e.name} size={24}/><div><div style={{fontWeight:600,fontSize:12}}>{e.name}</div><div style={{fontSize:10,color:C.muted}}>{e.id}</div></div></div></TD>
        <TD style={{fontSize:11}}>{e.nationality}</TD>
        <TD style={{fontSize:12}}>{sgd(p?.basic||basic+allow)}</TD>
        <TD style={{fontSize:12,color:C.teal}}>{sgd(p?.otPay||pendingOT)}</TD>
        <TD style={{fontWeight:600}}>{sgd(p?.gross||basic+allow+pendingOT)}</TD>
        <TD style={{color:C.danger,fontSize:12}}>{sgd(p?.cpfEmployee||cpf.employee)}</TD>
        <TD style={{color:C.warning,fontSize:12}}>{sgd(p?.cpfEmployer||cpf.employer)}</TD>
        <TD style={{fontSize:11}}>{sgd(p?.sdl||calcSDF(basic))}</TD>
        <TD style={{fontWeight:700,color:C.success}}>{sgd(p?.netPay||(basic+allow+pendingOT-cpf.employee))}</TD>
        <TD><Badge s={p?.status||"Draft"}/></TD>
        <TD><div style={{display:"flex",gap:4}}>
          <Btn v="ghost" sm onClick={()=>setShowSlip({emp:e,p})}>Slip</Btn>
          {!p&&<Btn v="success" sm onClick={()=>processOne(e)}>Process</Btn>}
          {p&&p.status==="Draft"&&<Btn v="primary" sm onClick={()=>onPublish(p.id)}>Publish</Btn>}
          {p&&<Btn v="danger" sm onClick={()=>setConfirmId(p.id)}>Del</Btn>}
        </div></TD>
      </TR>;})}/>
    {showSlip&&<Modal title="Payslip Preview" onClose={()=>setShowSlip(null)} width={440}><EmpPayslip empId={showSlip.emp.id} payroll={[...(showSlip.p?[{...showSlip.p,status:"Published"}]:[])]} employees={employees} lang={lang}/></Modal>}
    {confirmId&&<Confirm msg="Delete payroll record?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: SHIFT SCHEDULE
// ═══════════════════════════════════════════════════════════════════════════════
function AdminShift({shifts,employees,onAdd,onEdit,onDelete,lang}){
  const [showForm,setShowForm]=useState(false);
  const [editS,setEditS]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const [empF,setEmpF]=useState("");
  const [f,setF]=useState({empId:"",weekStart:"",shifts:[]});
  const DAYS_OF_WEEK=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
  function initShifts(weekStart){
    if(!weekStart)return[];
    const ws=new Date(weekStart);
    return DAYS_OF_WEEK.map((d,i)=>{const date=new Date(ws);date.setDate(ws.getDate()+(i===6?-1:i));return{date:date.toISOString().slice(0,10),type:"Day",start:"07:00",end:"17:00",site:""};});
  }
  function openNew(){setEditS(null);const today=new Date();const mon=new Date(today);mon.setDate(today.getDate()-today.getDay()+1);const ws=mon.toISOString().slice(0,10);setF({empId:"",weekStart:ws,shifts:initShifts(ws)});setShowForm(true);}
  function updateShift(i,k,v){setF(p=>({...p,shifts:p.shifts.map((s,x)=>x===i?{...s,[k]:v}:s)}));}
  function save(){if(!f.empId||!f.weekStart)return;editS?onEdit(editS.id,f):onAdd({...f,id:uid()});setShowForm(false);}
  const filtered=shifts.filter(s=>!empF||s.empId===empF);
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"shift")} actions={[<Btn onClick={openNew}>+ Assign Shift</Btn>]}/>
    <FilterBar><Sel label="" value={empF} onChange={setEmpF} options={[{v:"",l:"All Employees"},...employees.map(e=>({v:e.id,l:`${e.id} - ${e.name}`}))]}/></FilterBar>
    <TTable cols={["Employee","Week","Mon","Tue","Wed","Thu","Fri","Sat","Sun","Actions"]}
      rows={filtered.map(s=>{const emp=employees.find(e=>e.id===s.empId);const shiftC={Day:C.success,Night:C.purple,Off:C.muted,AL:C.warning};return <TR key={s.id}>
        <TD style={{fontWeight:600,fontSize:12}}>{emp?.name?.split(",")[0]||s.empId}</TD>
        <TD style={{fontSize:11}}>{fmtDate(s.weekStart)}</TD>
        {s.shifts.map((sh,i)=><TD key={i}><span style={{fontSize:10,fontWeight:700,color:shiftC[sh.type]||C.muted,background:shiftC[sh.type]+"22",borderRadius:4,padding:"2px 5px"}}>{sh.type}</span></TD>)}
        <TD><div style={{display:"flex",gap:4}}><Btn v="outline" sm onClick={()=>{setEditS(s);setF({...s});setShowForm(true);}}>Edit</Btn><Btn v="danger" sm onClick={()=>setConfirmId(s.id)}>Del</Btn></div></TD>
      </TR>;})}/>
    {showForm&&<Modal title={editS?"Edit Shift Schedule":"Assign Shift"} onClose={()=>setShowForm(false)} width={760}>
      <div style={{display:"flex",gap:12,marginBottom:16,flexWrap:"wrap"}}>
        <div style={{flex:1}}><Sel label="Employee" value={f.empId} onChange={v=>setF(p=>({...p,empId:v}))} options={employees.map(e=>({v:e.id,l:`${e.id} - ${e.name}`}))}/></div>
        <div style={{flex:1}}><Inp label="Week Starting (Monday)" type="date" value={f.weekStart} onChange={v=>setF(p=>({...p,weekStart:v,shifts:p.shifts.length?p.shifts:initShifts(v)}))}/></div>
      </div>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr>{["Day","Date","Type","Start","End","Site"].map(h=><th key={h} style={{padding:"6px 8px",textAlign:"left",fontWeight:700,color:C.muted,borderBottom:`1px solid ${C.border}`,fontSize:11}}>{h}</th>)}</tr></thead>
        <tbody>{f.shifts.map((sh,i)=><tr key={i} style={{borderBottom:`1px solid ${C.border}`}}>
          <td style={{padding:"6px 8px",fontWeight:700}}>{new Date(sh.date).toLocaleDateString("en-SG",{weekday:"short"})}</td>
          <td style={{padding:"6px 8px",fontSize:11,color:C.muted}}>{fmtDate(sh.date)}</td>
          <td style={{padding:"6px 8px"}}><select value={sh.type} onChange={e=>updateShift(i,"type",e.target.value)} style={{...IS,width:70,padding:"4px 6px",fontSize:12}}>{["Day","Night","Off","AL"].map(o=><option key={o}>{o}</option>)}</select></td>
          <td style={{padding:"6px 8px"}}><input type="time" value={sh.start||""} onChange={e=>updateShift(i,"start",e.target.value)} disabled={sh.type==="Off"||sh.type==="AL"} style={{...IS,width:90,padding:"4px 6px",fontSize:12}}/></td>
          <td style={{padding:"6px 8px"}}><input type="time" value={sh.end||""} onChange={e=>updateShift(i,"end",e.target.value)} disabled={sh.type==="Off"||sh.type==="AL"} style={{...IS,width:90,padding:"4px 6px",fontSize:12}}/></td>
          <td style={{padding:"6px 8px"}}><input value={sh.site||""} onChange={e=>updateShift(i,"site",e.target.value)} disabled={sh.type==="Off"||sh.type==="AL"} placeholder="Site name" style={{...IS,width:120,padding:"4px 6px",fontSize:12}}/></td>
        </tr>)}</tbody>
      </table></div>
      <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:14}}><Btn v="ghost" onClick={()=>setShowForm(false)}>{t(lang,"cancel")}</Btn><Btn onClick={save}>{t(lang,"save")}</Btn></div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete shift schedule?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
function AdminDashboard({employees,leaves,claims,attendance,announcements,onNavigate,onLeaveAction,session,lang}){
  const today=new Date();
  const pendingLeaves=leaves.filter(l=>l.status==="Pending Supervisor"||l.status==="Pending HR");
  const pendingClaims=claims.filter(c=>c.status==="Pending Supervisor"||c.status==="Pending HR");
  const alerts=[];employees.forEach(e=>{const ep=daysUntil(e.epExpiry);if(ep!==null&&ep<=90&&ep>=0)alerts.push({name:e.name,type:"EP/SP",days:ep,color:ep<=30?C.danger:C.warning});const pp=daysUntil(e.passportExpiry);if(pp!==null&&pp<=90&&pp>=0)alerts.push({name:e.name,type:"Passport",days:pp,color:pp<=60?C.danger:C.warning});});
  return <div style={{display:"flex",flexDirection:"column",gap:18}}>
    <div><h1 style={{fontSize:20,fontWeight:800,color:C.text,margin:0}}>Dashboard</h1><p style={{color:C.muted,margin:"3px 0 0",fontSize:13}}>{today.toLocaleDateString("en-SG",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</p></div>
    <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
      <StatCard label={t(lang,"totalEmployees")} value={employees.length} sub={`${employees.filter(e=>e.status==="Active").length} ${t(lang,"active")}`} icon="E" color={C.accent}/>
      <StatCard label="Pending Leaves" value={pendingLeaves.length} icon="L" color={C.warning}/>
      <StatCard label="Pending Claims" value={pendingClaims.length} icon="C" color={C.purple}/>
      <StatCard label="Doc Alerts" value={alerts.length} icon="!" color={C.danger}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><SecTitle>Pending Approvals</SecTitle><Btn v="outline" sm onClick={()=>onNavigate("leaves")}>View All</Btn></div>
        {pendingLeaves.length===0&&<p style={{color:C.muted,fontSize:13}}>All clear.</p>}
        {pendingLeaves.slice(0,4).map(l=>{const emp=employees.find(e=>e.id===l.empId);return <div key={l.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${C.border}`}}>
          <Avatar name={emp?.name||"?"} size={30}/><div style={{flex:1}}><div style={{fontWeight:600,fontSize:12}}>{emp?.name}</div><div style={{fontSize:11,color:C.muted}}>{l.type} · {fmtDate(l.from)}-{fmtDate(l.to)}</div><Badge s={l.status}/></div>
          <Btn v="success" sm onClick={()=>onLeaveAction(l.id,l.status==="Pending Supervisor"?"Pending HR":"Approved",session.name,"")}>✓</Btn>
          <Btn v="danger" sm onClick={()=>onLeaveAction(l.id,"Rejected",session.name,"")}>✕</Btn>
        </div>;})}
      </Card>
      <Card>
        <SecTitle>Document Expiry Alerts</SecTitle>
        {alerts.length===0&&<p style={{color:C.muted,fontSize:13}}>No alerts within 90 days.</p>}
        {alerts.slice(0,6).map((a,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:a.color,flexShrink:0}}/>
          <div style={{flex:1}}><span style={{fontWeight:600,fontSize:12}}>{a.name}</span><span style={{color:C.muted,fontSize:11}}> — {a.type}</span></div>
          <span style={{fontSize:12,fontWeight:700,color:a.color}}>{a.days===0?"TODAY":a.days+"d"}</span>
        </div>)}
      </Card>
      <Card>
        <SecTitle>Announcements</SecTitle>
        {announcements.slice(0,3).map(a=><div key={a.id} style={{borderLeft:`3px solid ${C.accent}`,paddingLeft:10,marginBottom:10}}><div style={{fontWeight:600,fontSize:13}}>{a.title}</div><div style={{fontSize:11,color:C.muted}}>{fmtDate(a.date)}</div></div>)}
      </Card>
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// LOGIN
// ═══════════════════════════════════════════════════════════════════════════════
function Login({users,onLogin,lang,setLang}){
  const [user,setUser]=useState("");const [pass,setPass]=useState("");const [err,setErr]=useState("");
  function submit(){const found=users.find(x=>x.username.toLowerCase()===user.trim().toLowerCase()&&x.password===pass);if(!found){setErr("Invalid credentials.");return;}if(found.active===false){setErr("Account disabled. Contact admin.");return;}onLogin(found);}
  return <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#1A2340 0%,#2563EB 100%)",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
    <div style={{background:"#fff",borderRadius:16,padding:"40px",width:380,boxShadow:"0 24px 64px rgba(0,0,0,0.25)"}}>
      <div style={{textAlign:"center",marginBottom:28}}>
        <div style={{width:56,height:56,background:C.accent,borderRadius:14,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:800,marginBottom:12,color:"#fff"}}>C.HR</div>
        <div style={{fontSize:22,fontWeight:800,color:C.text}}>Custera<span style={{color:C.accent}}>.HR</span></div>
        <div style={{color:C.muted,fontSize:12,marginTop:4}}>Singapore Construction HR System</div>
        <div style={{display:"flex",justifyContent:"center",gap:8,marginTop:10}}>
          <button onClick={()=>setLang("en")} style={{background:lang==="en"?C.accent:"#f1f5f9",color:lang==="en"?"#fff":C.muted,border:"none",borderRadius:6,padding:"4px 12px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>EN</button>
          <button onClick={()=>setLang("zh")} style={{background:lang==="zh"?C.accent:"#f1f5f9",color:lang==="zh"?"#fff":C.muted,border:"none",borderRadius:6,padding:"4px 12px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>中文</button>
        </div>
      </div>
      {err&&<div style={{background:C.dangerL,color:C.dangerD,borderRadius:8,padding:"9px 14px",marginBottom:14,fontSize:13,fontWeight:600}}>{err}</div>}
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Inp label="Username / Email" value={user} onChange={setUser} placeholder="Enter username"/>
        <Inp label="Password" type="password" value={pass} onChange={setPass} placeholder="Enter password"/>
        <button onClick={submit} style={{background:C.accent,color:"#fff",border:"none",borderRadius:8,padding:12,fontWeight:700,fontSize:14,cursor:"pointer",marginTop:4,fontFamily:"inherit"}}>Sign In</button>
      </div>
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ROOT APP
// ═══════════════════════════════════════════════════════════════════════════════
const EMP_NAV=[{id:"dashboard",l:"dashboard",icon:"D"},{id:"attendance",l:"attendance",icon:"A"},{id:"leaves",l:"leaves",icon:"L"},{id:"claims",l:"claims",icon:"C"},{id:"payslip",l:"payslip",icon:"$"},{id:"memo",l:"memo",icon:"M"},{id:"policy",l:"policy",icon:"P"},{id:"training",l:"training",icon:"T"},{id:"feedback",l:"feedback",icon:"F"},{id:"appraisal",l:"appraisal",icon:"R"},{id:"calendar",l:"calendar",icon:"K"},{id:"shift",l:"shift",icon:"S"}];
const ADMIN_NAV=[{id:"dashboard",l:"dashboard",icon:"D"},{id:"employees",l:"employees",icon:"E"},{id:"users",l:"settings",icon:"U"},{id:"attendance",l:"attendance",icon:"A"},{id:"leaves",l:"leaves",icon:"L"},{id:"claims",l:"claims",icon:"C"},{id:"payroll",l:"payroll",icon:"$"},{id:"memo",l:"memo",icon:"M"},{id:"policy",l:"policy",icon:"P"},{id:"training",l:"training",icon:"T"},{id:"feedback",l:"feedback",icon:"F"},{id:"appraisal",l:"appraisal",icon:"R"},{id:"calendar",l:"calendar",icon:"K"},{id:"shift",l:"shift",icon:"S"}];

export default function App(){
  const [session,setSession]=useState(null);
  const [lang,setLang]=useState("en");
  const [page,setPage]=useState("dashboard");
  const [employees,setEmployees]=useState(INIT_EMPLOYEES);
  const [users,setUsers]=useState(INIT_USERS);
  const [leaves,setLeaves]=useState(INIT_LEAVES);
  const [leaveTypes,setLeaveTypes]=useState(SG_LEAVE_TYPES);
  const [claims,setClaims]=useState(INIT_CLAIMS);
  const [payroll,setPayroll]=useState(INIT_PAYROLL);
  const [memos,setMemos]=useState(INIT_MEMOS);
  const [policies,setPolicies]=useState(INIT_POLICIES);
  const [trainings,setTrainings]=useState(INIT_TRAININGS);
  const [feedbacks,setFeedbacks]=useState(INIT_FEEDBACKS);
  const [feedbackSections,setFeedbackSections]=useState(INIT_FEEDBACK_SECTIONS);
  const [appraisalForms,setAppraisalForms]=useState(INIT_APPRAISAL_FORMS);
  const [appraisalSubs,setAppraisalSubs]=useState(INIT_APPRAISAL_SUBMISSIONS);
  const [attendance,setAttendance]=useState(INIT_ATTENDANCE);
  const [shifts,setShifts]=useState(INIT_SHIFTS);
  const [calEvents,setCalEvents]=useState(INIT_CALENDAR_EVENTS);
  const [toast,setToast]=useState(null);

  function toast_(msg,type){setToast({msg,type:type||"success"});}
  const isAdmin=session&&(session.role==="admin"||session.role==="superadmin");
  const isSupervisor=session?.role==="supervisor";
  const isAdminOrSup=isAdmin||isSupervisor;

  if(!session)return <Login users={users} onLogin={u=>{setSession(u);setPage("dashboard");}} lang={lang} setLang={setLang}/>;

  const empId=session.empId;
  const NAV=isAdminOrSup?ADMIN_NAV:EMP_NAV;

  // CRUD helpers
  const crud=(setter,toastMsg)=>({
    add:(d)=>{setter(p=>[...p,{...d,id:d.id||uid()}]);toast_(toastMsg||(t(lang,"add")+" ok"));},
    edit:(id,d)=>{setter(p=>p.map(x=>x.id===id?{...x,...d,id}:x));toast_(t(lang,"save")+" ok");},
    del:(id)=>{setter(p=>p.filter(x=>x.id!==id));toast_("Deleted","error");},
  });

  const empCrud=crud(setEmployees,"Employee saved");
  const userCrud={
    add:(d)=>{setUsers(p=>[...p,{...d,id:d.id||"u"+uid()}]);toast_("Account created");},
    edit:(id,d)=>{setUsers(p=>p.map(x=>x.id===id?{...x,...d,id}:x));toast_("Account updated");},
    del:(id)=>{setUsers(p=>p.filter(x=>x.id!==id));toast_("Account deleted","error");},
    resetPw:(id,pw)=>{setUsers(p=>p.map(x=>x.id===id?{...x,password:pw}:x));toast_("Password reset");},
  };

  function addLeave(d){setLeaves(p=>[...p,{...d,id:uid()}]);toast_("Leave submitted");}
  function leaveAction(id,status,by,comment){setLeaves(p=>p.map(l=>l.id===id?{...l,status,supervisorComment:l.status==="Pending Supervisor"?comment:l.supervisorComment,hrComment:l.status==="Pending HR"?comment:l.hrComment}:l));toast_(status==="Approved"?"Approved":status==="Rejected"?"Rejected":"Updated",status==="Rejected"?"error":"success");}
  function delLeave(id){setLeaves(p=>p.filter(l=>l.id!==id));toast_("Deleted","error");}
  function addLT(d){setLeaveTypes(p=>[...p,d]);}function editLT(i,d){setLeaveTypes(p=>p.map((x,xi)=>xi===i?d:x));}function delLT(i){setLeaveTypes(p=>p.filter((_,xi)=>xi!==i));}

  function addClaim(d){setClaims(p=>[...p,{...d,id:uid()}]);toast_("Claim submitted");}
  function claimAction(id,status,by,comment){setClaims(p=>p.map(c=>c.id===id?{...c,status,supervisorComment:c.status==="Pending Supervisor"?by:c.supervisorComment,hrComment:c.status==="Pending HR"?by:c.hrComment}:c));toast_(status,"success");}
  function delClaim(id){setClaims(p=>p.filter(c=>c.id!==id));toast_("Deleted","error");}

  function clockAction(rec){
    const existing=attendance.find(a=>a.empId===rec.empId&&a.date===rec.date);
    if(existing){setAttendance(p=>p.map(a=>a.id===existing.id?{...a,...rec,id:a.id}:a));}
    else{setAttendance(p=>[...p,{...rec,id:uid()}]);}
    toast_("Attendance recorded");
  }

  function processPayroll(d){setPayroll(p=>[...p,d]);toast_("Payroll processed");}
  function delPayroll(id){setPayroll(p=>p.filter(x=>x.id!==id));toast_("Deleted","error");}
  function publishPayslip(id){setPayroll(p=>p.map(x=>x.id===id?{...x,status:"Published",publishedOn:todayStr()}:x));toast_("Payslip published to employee");}

  const memoCrud=crud(setMemos,"Announcement saved");
  const policyCrud=crud(setPolicies,"Policy saved");
  const trainingCrud=crud(setTrainings,"Training saved");
  const shiftCrud=crud(setShifts,"Shift saved");
  const calCrud=crud(setCalEvents,"Event saved");
  const appraisalFormCrud=crud(setAppraisalForms,"Form saved");

  function readMemo(id){setMemos(p=>p.map(m=>m.id===id?{...m,readBy:[...m.readBy,empId]}:m));}
  function replyFeedback(id,reply){setFeedbacks(p=>p.map(f=>f.id===id?{...f,adminReply:reply,status:"Closed"}:f));toast_("Reply sent");}
  function addFeedbackSection(d){setFeedbackSections(p=>[...p,{...d,id:uid()}]);toast_("Category added");}
  function editFeedbackSection(id,d){setFeedbackSections(p=>p.map(s=>s.id===id?{...s,...d,id}:s));toast_("Category updated");}
  function delFeedbackSection(id){setFeedbackSections(p=>p.filter(s=>s.id!==id));toast_("Deleted","error");}
  function addFeedback(d){setFeedbacks(p=>[...p,{...d,id:uid()}]);toast_("Feedback submitted");}
  function saveAppraisalSub(d){setAppraisalSubs(p=>{const ex=p.find(s=>s.id===d.id);return ex?p.map(s=>s.id===d.id?d:s):[...p,d]});toast_("Appraisal saved");}
  function editAppraisalForm(id,d){setAppraisalForms(p=>p.map(f=>f.id===id?{...f,...d,id}:f));toast_("Form updated");}

  const unreadMemos=memos.filter(m=>!m.readBy.includes(empId));
  const showMemoPopup=!isAdminOrSup&&unreadMemos.length>0;

  function renderPage(){
    if(!isAdminOrSup){
      switch(page){
        case "dashboard":return <EmpDashboard session={session} employees={employees} leaves={leaves} claims={claims} memos={memos} shifts={shifts} payroll={payroll} onNavigate={setPage} lang={lang}/>;
        case "attendance":return <EmpAttendance empId={empId} attendance={attendance} onClock={clockAction} lang={lang}/>;
        case "leaves":return <EmpLeave empId={empId} leaves={leaves} leaveTypes={leaveTypes} onAdd={addLeave} lang={lang}/>;
        case "claims":return <EmpClaims empId={empId} claims={claims} onAdd={addClaim} lang={lang}/>;
        case "payslip":return <EmpPayslip empId={empId} payroll={payroll} employees={employees} lang={lang}/>;
        case "memo":return <EmpMemo memos={memos} empId={empId} lang={lang} onRead={readMemo}/>;
        case "policy":return <EmpPolicy policies={policies} lang={lang}/>;
        case "training":return <EmpTraining trainings={trainings} empId={empId} lang={lang}/>;
        case "feedback":return <EmpFeedback feedbacks={feedbacks} sections={feedbackSections} empId={empId} onAdd={addFeedback} lang={lang}/>;
        case "appraisal":return <EmpAppraisal forms={appraisalForms} submissions={appraisalSubs} empId={empId} onSave={saveAppraisalSub} lang={lang}/>;
        case "calendar":return <EmpCalendar events={calEvents} leaves={leaves} employees={employees} empId={empId} lang={lang}/>;
        case "shift":return <EmpShift shifts={shifts} empId={empId} lang={lang}/>;
        default:return null;
      }
    }
    switch(page){
      case "dashboard":return <AdminDashboard employees={employees} leaves={leaves} claims={claims} attendance={attendance} announcements={memos} onNavigate={setPage} onLeaveAction={leaveAction} session={session} lang={lang}/>;
      case "employees":return <EmpList employees={employees} leaves={leaves} onAdd={empCrud.add} onEdit={empCrud.edit} onDelete={empCrud.del} lang={lang}/>;
      case "users":return <UserMgmt users={users} employees={employees} onAdd={userCrud.add} onEdit={userCrud.edit} onDelete={userCrud.del} onResetPw={userCrud.resetPw} lang={lang}/>;
      case "attendance":return <AdminAttendance attendance={attendance} employees={employees} lang={lang}/>;
      case "leaves":return <AdminLeave leaves={leaves} employees={employees} leaveTypes={leaveTypes} users={users} onAction={leaveAction} onAdd={addLeave} onDelete={delLeave} onAddType={addLT} onEditType={editLT} onDeleteType={delLT} session={session} lang={lang}/>;
      case "claims":return <AdminClaims claims={claims} employees={employees} users={users} session={session} onAction={claimAction} onDelete={delClaim} lang={lang}/>;
      case "payroll":return <AdminPayroll employees={employees} payroll={payroll} claims={claims} onProcess={processPayroll} onDelete={delPayroll} onPublish={publishPayslip} lang={lang}/>;
      case "memo":return <AdminMemo memos={memos} onAdd={memoCrud.add} onEdit={memoCrud.edit} onDelete={memoCrud.del} session={session} lang={lang}/>;
      case "policy":return <AdminPolicy policies={policies} onAdd={policyCrud.add} onEdit={policyCrud.edit} onDelete={policyCrud.del} session={session} lang={lang}/>;
      case "training":return <AdminTraining trainings={trainings} employees={employees} onAdd={trainingCrud.add} onEdit={trainingCrud.edit} onDelete={trainingCrud.del} session={session} lang={lang}/>;
      case "feedback":return <AdminFeedback feedbacks={feedbacks} sections={feedbackSections} employees={employees} onReply={replyFeedback} onDeleteFeedback={(id)=>setFeedbacks(p=>p.filter(f=>f.id!==id))} onAddSection={addFeedbackSection} onEditSection={editFeedbackSection} onDeleteSection={delFeedbackSection} lang={lang}/>;
      case "appraisal":return <AdminAppraisal forms={appraisalForms} submissions={appraisalSubs} employees={employees} onAddForm={appraisalFormCrud.add} onEditForm={editAppraisalForm} onDeleteForm={appraisalFormCrud.del} onSaveSubmission={saveAppraisalSub} session={session} lang={lang}/>;
      case "calendar":return <AdminCalendar events={calEvents} leaves={leaves} employees={employees} onAdd={calCrud.add} onEdit={calCrud.edit} onDelete={calCrud.del} session={session} lang={lang}/>;
      case "shift":return <AdminShift shifts={shifts} employees={employees} onAdd={shiftCrud.add} onEdit={shiftCrud.edit} onDelete={shiftCrud.del} lang={lang}/>;
      default:return null;
    }
  }

  return <div style={{display:"flex",height:"100vh",fontFamily:"'Inter',-apple-system,BlinkMacSystemFont,sans-serif",background:C.bg,overflow:"hidden"}}>
    {/* Sidebar */}
    <aside style={{width:190,background:C.sidebar,display:"flex",flexDirection:"column",flexShrink:0,overflowY:"auto"}}>
      <div style={{padding:"16px 14px 12px",borderBottom:"1px solid rgba(255,255,255,0.08)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:32,height:32,background:C.accent,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"#fff",flexShrink:0}}>C.HR</div>
          <div><div style={{color:"#fff",fontWeight:800,fontSize:13,lineHeight:1.1}}>Custera<span style={{color:"#60A5FA"}}>.HR</span></div><div style={{color:"rgba(255,255,255,0.4)",fontSize:9,marginTop:1}}>SG Construction</div></div>
        </div>
        <div style={{display:"flex",gap:4,marginTop:10}}>
          {["en","zh"].map(l=><button key={l} onClick={()=>setLang(l)} style={{flex:1,background:lang===l?C.accent:"rgba(255,255,255,0.08)",color:lang===l?"#fff":"rgba(255,255,255,0.5)",border:"none",borderRadius:5,padding:"3px 0",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{l==="en"?"EN":"中文"}</button>)}
        </div>
      </div>
      <nav style={{flex:1,padding:"8px 6px"}}>
        {NAV.map(n=>{const active=page===n.id;return <button key={n.id} onClick={()=>setPage(n.id)}
          style={{display:"flex",alignItems:"center",gap:8,width:"100%",padding:"8px 9px",borderRadius:7,border:"none",background:active?C.sidebarA:"transparent",color:active?"#93C5FD":C.sidebarT,fontWeight:active?700:500,fontSize:12,cursor:"pointer",marginBottom:1,textAlign:"left",fontFamily:"inherit"}}
          onMouseEnter={e=>{if(!active)e.currentTarget.style.background="rgba(255,255,255,0.07)";}}
          onMouseLeave={e=>{if(!active)e.currentTarget.style.background="transparent";}}>
          <span style={{width:18,height:18,background:"rgba(255,255,255,0.1)",borderRadius:4,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:800,flexShrink:0}}>{n.icon}</span>
          {t(lang,n.l)}
          {n.id==="memo"&&!isAdminOrSup&&unreadMemos.length>0&&<span style={{marginLeft:"auto",background:C.danger,color:"#fff",fontSize:9,fontWeight:800,borderRadius:99,padding:"1px 5px",flexShrink:0}}>{unreadMemos.length}</span>}
        </button>;})}
      </nav>
      <div style={{padding:"10px",borderTop:"1px solid rgba(255,255,255,0.08)"}}>
        <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}>
          <Avatar name={session.name} size={26}/>
          <div style={{flex:1,overflow:"hidden"}}><div style={{color:"#fff",fontWeight:600,fontSize:11,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{session.name}</div><div style={{color:"rgba(255,255,255,0.4)",fontSize:9,textTransform:"capitalize"}}>{session.role==="superadmin"?"Super Admin":session.role}</div></div>
        </div>
        <button onClick={()=>setSession(null)} style={{width:"100%",background:"rgba(255,255,255,0.07)",border:"none",borderRadius:6,padding:"5px",color:"rgba(255,255,255,0.5)",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>{t(lang,"signOut")}</button>
      </div>
    </aside>
    <main style={{flex:1,overflowY:"auto",padding:"20px 24px"}}>{renderPage()}</main>
    {toast&&<Toast msg={toast.msg} type={toast.type} onDone={()=>setToast(null)}/>}
    {showMemoPopup&&<MemoPopup memos={memos} empId={empId} lang={lang} onRead={readMemo} onClose={()=>readMemo(unreadMemos[0].id)}/>}
  </div>;
}

// Placeholder employee list for admin (reuse from previous version)
function EmpList({employees,leaves,onAdd,onEdit,onDelete,lang}){
  const [search,setSearch]=useState("");
  const [confirmId,setConfirmId]=useState(null);
  const [editE,setEditE]=useState(null);
  const [showAdd,setShowAdd]=useState(false);
  const filtered=employees.filter(e=>e.name.toLowerCase().includes(search.toLowerCase())||e.id.toLowerCase().includes(search.toLowerCase())||(e.fin||"").toLowerCase().includes(search.toLowerCase()));
  const [f,setF]=useState({});const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  function openEdit(e){setEditE(e);setF({...e});setShowAdd(true);}
  function save(){if(!f.name)return;editE?onEdit(editE.id,f):onAdd({...f,id:f.id||"C"+String(employees.length+1).padStart(3,"0")});setShowAdd(false);setEditE(null);}
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"employees")} sub={`${filtered.length} records`} actions={[<Btn onClick={()=>{setEditE(null);setF({status:"Active",annualLeave:14,company:"STEC"});setShowAdd(true);}}>+ Add</Btn>,<Btn v="ghost" onClick={()=>exportXLS(employees.map(e=>({ID:e.id,Name:e.name,"FIN/NRIC":e.fin,Company:e.company,Dept:e.department,Position:e.position,Nationality:e.nationality,"Work Pass":e.workPass,DOB:fmtDate(e.dob),Mobile:e.mobile,Address:e.address,Status:e.status})),"Employees","staff_list.xlsx")}>{t(lang,"export")}</Btn>]}/>
    <FilterBar><SearchInp value={search} onChange={setSearch} placeholder="Search name, ID, FIN..."/></FilterBar>
    <TTable cols={["ID","Employee","Company","Dept","Position","Work Pass","Status","Actions"]}
      rows={filtered.map(e=><TR key={e.id}>
        <TD style={{fontWeight:700,color:C.accent,fontSize:12}}>{e.id}</TD>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={e.name} size={28}/><div><div style={{fontWeight:600,fontSize:12}}>{e.name}</div><div style={{fontSize:11,color:C.muted}}>{e.mobile}</div></div></div></TD>
        <TD style={{fontSize:12}}>{e.company}</TD><TD style={{fontSize:12}}>{e.department}</TD><TD style={{fontSize:12}}>{e.position}</TD>
        <TD><Badge s={e.workPass||"—"}/></TD><TD><Badge s={e.status}/></TD>
        <TD><div style={{display:"flex",gap:4}}><Btn v="outline" sm onClick={()=>openEdit(e)}>Edit</Btn><Btn v="danger" sm onClick={()=>setConfirmId(e.id)}>Del</Btn></div></TD>
      </TR>)}/>
    {showAdd&&<Modal title={editE?"Edit Employee":"Add Employee"} onClose={()=>setShowAdd(false)} width={640}>
      <Grid cols={2}>
        {[["Employee ID","id","text"],["Full Name","name","text"],["FIN/NRIC","fin","text"],["Department","department","text"],["Company","company","text"],["Position","position","text"],["Date Joined","dateJoined","date"],["Date of Birth","dob","date"],["Nationality","nationality","text"],["Work Pass","workPass","text"],["Mobile","mobile","text"],["Basic Salary","basicSalary","number"],["Allowance","allowance","number"],["Status","status","text"]].map(([l,k,tp])=><Inp key={k} label={l} type={tp} value={f[k]||""} onChange={v=>setFk(k,v)} readOnly={k==="id"&&!!editE}/>)}
      </Grid>
      <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:12}}><Btn v="ghost" onClick={()=>setShowAdd(false)}>Cancel</Btn><Btn onClick={save}>Save</Btn></div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete employee?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// Employee self-dashboard
function EmpDashboard({session,employees,leaves,claims,memos,shifts,payroll,onNavigate,lang}){
  const empId=session.empId;
  const myLeaves=leaves.filter(l=>l.empId===empId);
  const myClaims=claims.filter(c=>c.empId===empId);
  const myShift=shifts.filter(s=>s.empId===empId).sort((a,b)=>b.weekStart.localeCompare(a.weekStart))[0];
  const latestPayslip=payroll.filter(p=>p.empId===empId&&p.status==="Published").sort((a,b)=>b.month.localeCompare(a.month))[0];
  const emp=employees.find(e=>e.id===empId);
  const unreadMemos=memos.filter(m=>!m.readBy.includes(empId));
  const pendingLeave=myLeaves.filter(l=>l.status==="Pending Supervisor"||l.status==="Pending HR").length;
  const alUsed=myLeaves.filter(l=>l.status==="Approved"&&l.type==="Annual Leave").reduce((s,l)=>s+l.days,0);
  const alEnt=emp?.annualLeave||14;
  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    <div style={{background:`linear-gradient(135deg,${C.sidebar},${C.accent})`,borderRadius:14,padding:"20px 24px",color:"#fff"}}>
      <div style={{display:"flex",gap:14,alignItems:"center"}}>
        <Avatar name={session.name} size={52}/>
        <div><div style={{fontSize:18,fontWeight:800}}>{session.name}</div><div style={{fontSize:13,opacity:0.8}}>{emp?.position||""} · {emp?.department||""}</div><div style={{fontSize:11,opacity:0.6,marginTop:2}}>ID: {empId} · {fmtDate(todayStr())}</div></div>
      </div>
    </div>
    <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
      <StatCard label="AL Balance" value={`${alEnt-alUsed}d`} sub={`of ${alEnt} days`} icon="L" color={C.success}/>
      <StatCard label="Pending Leaves" value={pendingLeave} icon="P" color={C.warning}/>
      <StatCard label="Pending Claims" value={myClaims.filter(c=>c.status==="Pending Supervisor"||c.status==="Pending HR").length} icon="C" color={C.purple}/>
      <StatCard label="Unread Memos" value={unreadMemos.length} icon="M" color={unreadMemos.length>0?C.danger:C.success}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <Card style={{cursor:"pointer"}} onClick={()=>onNavigate("shift")}>
        <SecTitle>This Week's Shift</SecTitle>
        {myShift?<div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{myShift.shifts.map(s=>{const shC={Day:C.success,Night:C.purple,Off:C.muted,AL:C.warning};return <div key={s.date} style={{background:shC[s.type]+"22",borderRadius:6,padding:"4px 8px",fontSize:11,fontWeight:700,color:shC[s.type]}}>{new Date(s.date).toLocaleDateString("en-SG",{weekday:"short"})} {s.type}</div>;})}</div>:<p style={{color:C.muted,fontSize:13}}>No shift scheduled.</p>}
      </Card>
      <Card style={{cursor:"pointer"}} onClick={()=>onNavigate("payslip")}>
        <SecTitle>Latest Payslip</SecTitle>
        {latestPayslip?<><div style={{fontSize:14,fontWeight:700}}>{latestPayslip.month}</div><div style={{fontSize:20,fontWeight:800,color:C.success}}>{sgd(latestPayslip.netPay)}</div><div style={{fontSize:11,color:C.muted}}>Net Pay</div></>:<p style={{color:C.muted,fontSize:13}}>No payslip yet.</p>}
      </Card>
      {unreadMemos.length>0&&<Card style={{gridColumn:"span 2",cursor:"pointer",borderLeft:`4px solid ${C.danger}`}} onClick={()=>onNavigate("memo")}>
        <SecTitle>Unread Announcements</SecTitle>
        {unreadMemos.slice(0,2).map(m=><div key={m.id} style={{marginBottom:8}}><strong style={{fontSize:13}}>{lang==="zh"&&m.titleZh?m.titleZh:m.title}</strong><Badge s={m.tag}/><div style={{fontSize:12,color:C.muted,marginTop:2}}>{fmtDate(m.date)}</div></div>)}
      </Card>}
    </div>
  </div>;
}
