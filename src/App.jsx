import React, { useState, useEffect, useRef } from "react";
import { initializeApp } from "firebase/app";
import { getFirestore, collection, onSnapshot, addDoc, updateDoc, deleteDoc, doc, setDoc, getDocs, writeBatch } from "firebase/firestore";

const FB_CONFIG={apiKey:"AIzaSyAZ4tG7FDPHVN9O4HC5VEIjTqoJheMH6b8",authDomain:"custera-hr.firebaseapp.com",projectId:"custera-hr",storageBucket:"custera-hr.firebasestorage.app",messagingSenderId:"459279090099",appId:"1:459279090099:web:6e85fb5c6d4236efd58bc6"};
const fbApp=initializeApp(FB_CONFIG);
const db=getFirestore(fbApp);
import * as XLSX from "xlsx";

// ─── THEME ───────────────────────────────────────────────────────────────────
// ─── MODERN THEME ─────────────────────────────────────────────────────────────
const C = {
  sidebar:"#0F172A", sidebarA:"rgba(99,102,241,0.18)", sidebarT:"rgba(148,163,184,0.85)",
  accent:"#6366F1", accentL:"#EEF2FF", accentT:"#4338CA",
  bg:"#F1F5F9", card:"#FFFFFF", text:"#0F172A", muted:"#64748B", border:"#E2E8F0",
  success:"#10B981", successL:"#D1FAE5", successD:"#065F46",
  warning:"#F59E0B", warningL:"#FEF3C7", warningD:"#92400E",
  danger:"#EF4444", dangerL:"#FEE2E2", dangerD:"#991B1B",
  purple:"#8B5CF6", purpleL:"#EDE9FE",
  orange:"#F97316", orangeL:"#FFEDD5",
  teal:"#14B8A6", tealL:"#CCFBF1",
  pink:"#EC4899", pinkL:"#FCE7F3",
  navy:"#1E3A5F",
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
  { id:"u1", username:"admin", password:"admin123", role:"admin", name:"HR Admin", empId:"C010", dept:"HR & Admin.", reportingOfficerId:"", active:true },
  { id:"u2", username:"hrsuper", password:"super123", role:"supervisor", name:"GOH Siew Ling", empId:"C006", dept:"HR & Admin.", reportingOfficerId:"u1", active:true },
  { id:"u3", username:"staff1", password:"staff123", role:"employee", name:"TAN Wei Jie", empId:"C001", dept:"Management", reportingOfficerId:"u1", active:true },
  { id:"u4", username:"staff2", password:"staff234", role:"employee", name:"CHEN Jun Hao", empId:"C002", dept:"Engineering & Operations", reportingOfficerId:"u2", active:true },
  { id:"u5", username:"staff3", password:"staff345", role:"employee", name:"LIM Wee Kiat", empId:"C004", dept:"Business Development", reportingOfficerId:"u2", active:true },
];const INIT_EMPLOYEES = [
  {id:"C001",name:"TAN Wei Jie",fin:"S7812340A",company:"STEC",position:"Chief Executive Officer",department:"Management",dateJoined:"2011-05-01",gender:"M",dob:"1973-06-15",nationality:"Singaporean",workPass:"SC",epSpNo:"",workPassIssueDate:"",workPassExpiryDate:"",passportNumber:"K1234567A",passportIssueDate:"2020-01-15",passportExpiryDate:"2030-01-15",probationEndDate:"",qualification:"Bachelor of Civil Engineering, NUS",mobile:"9111 2222",residenceAddress:"Blk 123 Woodlands Drive 14, #05-01, Singapore 730123",personalEmail:"tanweijie@gmail.com",workEmail:"weijie@stec.com.sg",bankName:"DBS",bankAccount:"012-3-456789",emergencyContactName:"LIM Mei Hua",emergencyContactPhone:"9333 4444",emergencyContactRelation:"Spouse",phoneAllowance:0,housingAllowance:0,airfareAllowance:0,otherAllowance:0,basicSalary:0,allowance:0,annualLeave:14,status:"Active",site:"HQ",lastWorkingDay:""},
  {id:"C002",name:"CHEN Jun Hao",fin:"M4623450U",company:"Custera",position:"Chief Engineering & Operations Officer",department:"Engineering & Operations",dateJoined:"2026-01-27",gender:"M",dob:"1982-08-20",nationality:"Chinese",workPass:"EP",epSpNo:"M4623450U",workPassIssueDate:"2026-01-31",workPassExpiryDate:"2028-01-30",passportNumber:"PE3312345",passportIssueDate:"2022-03-06",passportExpiryDate:"2032-03-06",probationEndDate:"2026-04-27",qualification:"Master of Structural Engineering",mobile:"8982 3456",residenceAddress:"Blk 80 Compassvale Bow #08-40, Singapore 544570",personalEmail:"chenjh82@gmail.com",workEmail:"jh.chen@custera.com.sg",bankName:"DBS",bankAccount:"272-9-123456",emergencyContactName:"WANG Fang",emergencyContactPhone:"8234 5678",emergencyContactRelation:"Spouse",phoneAllowance:100,housingAllowance:800,airfareAllowance:1200,otherAllowance:100,basicSalary:8400,allowance:2200,annualLeave:14,status:"Active",site:"Tuas Site",lastWorkingDay:""},
  {id:"C004",name:"LIM Wee Kiat",fin:"S9779500I",company:"UTEC",position:"Business Development Engineer",department:"Business Development",dateJoined:"2022-02-01",gender:"M",dob:"1997-04-08",nationality:"Malaysian",workPass:"SPR",epSpNo:"",workPassIssueDate:"",workPassExpiryDate:"",passportNumber:"K5498396",passportIssueDate:"2021-12-06",passportExpiryDate:"2031-12-06",probationEndDate:"2022-08-01",qualification:"Bachelor of Mechanical Engineering, UTM",mobile:"8542 1234",residenceAddress:"APT BLK 280B Sengkang East Ave #15-629, Singapore 542280",personalEmail:"limwk97@gmail.com",workEmail:"weekiat@utec.com.sg",bankName:"DBS",bankAccount:"271-0-614500",emergencyContactName:"LIM Ah Kow",emergencyContactPhone:"9183 7826",emergencyContactRelation:"Father",phoneAllowance:50,housingAllowance:0,airfareAllowance:0,otherAllowance:50,basicSalary:4500,allowance:500,annualLeave:14,status:"Active",site:"Jurong Site",lastWorkingDay:""},
  {id:"C005",name:"RAHMAN Bin Aziz",fin:"S9686150F",company:"UTEC",position:"Site Engineer",department:"Engineering & Operations",dateJoined:"2022-04-01",gender:"M",dob:"1996-09-28",nationality:"Malaysian",workPass:"SPR",epSpNo:"",workPassIssueDate:"",workPassExpiryDate:"",passportNumber:"A5541790",passportIssueDate:"2021-08-20",passportExpiryDate:"2031-08-20",probationEndDate:"2022-10-01",qualification:"Diploma in Civil Engineering",mobile:"9340 2222",residenceAddress:"Blk 13 Cantonment Close #26-29, Singapore 080013",personalEmail:"rahmanaziz96@gmail.com",workEmail:"rahman@utec.com.sg",bankName:"DBS",bankAccount:"271-1-397500",emergencyContactName:"SITI Binte Aziz",emergencyContactPhone:"8540 4034",emergencyContactRelation:"Sister",phoneAllowance:50,housingAllowance:0,airfareAllowance:0,otherAllowance:0,basicSalary:4200,allowance:400,annualLeave:14,status:"Active",site:"Tuas Site",lastWorkingDay:""},
  {id:"C006",name:"GOH Siew Ling",fin:"S7579860D",company:"STEC",position:"Assistant Admin Manager",department:"HR & Admin.",dateJoined:"2025-10-15",gender:"F",dob:"1975-03-28",nationality:"Singaporean",workPass:"SC",epSpNo:"",workPassIssueDate:"",workPassExpiryDate:"",passportNumber:"K7654321B",passportIssueDate:"2021-05-10",passportExpiryDate:"2031-05-10",probationEndDate:"2026-04-15",qualification:"Diploma in Business Administration",mobile:"9748 5678",residenceAddress:"53 Serangoon Terrace, Singapore 535787",personalEmail:"gohsiewling@gmail.com",workEmail:"siewling@stec.com.sg",bankName:"OCBC",bankAccount:"512-3-456789",emergencyContactName:"GOH Ah Beng",emergencyContactPhone:"9234 5678",emergencyContactRelation:"Spouse",phoneAllowance:50,housingAllowance:0,airfareAllowance:0,otherAllowance:0,basicSalary:4800,allowance:100,annualLeave:14,status:"Active",site:"HQ",lastWorkingDay:""},
  {id:"C010",name:"TAN Li Fen",fin:"T0112340H",company:"STEC",position:"HR cum Admin Executive",department:"HR & Admin.",dateJoined:"2026-01-19",gender:"F",dob:"2001-03-15",nationality:"Singaporean",workPass:"SC",epSpNo:"",workPassIssueDate:"",workPassExpiryDate:"",passportNumber:"K4382218H",passportIssueDate:"2023-09-10",passportExpiryDate:"2033-09-10",probationEndDate:"2026-07-19",qualification:"Diploma in Human Resource Management",mobile:"9053 1234",residenceAddress:"201 Marsiling Drive #07-112, Singapore 730201",personalEmail:"tanlf01@gmail.com",workEmail:"lifen@stec.com.sg",bankName:"UOB",bankAccount:"453-3-006350",emergencyContactName:"TAN Ah Kow",emergencyContactPhone:"8286 1234",emergencyContactRelation:"Father",phoneAllowance:50,housingAllowance:0,airfareAllowance:0,otherAllowance:0,basicSalary:3500,allowance:50,annualLeave:14,status:"Active",site:"HQ",lastWorkingDay:""},
  {id:"C011",name:"CHONG Ah Mui",fin:"S1218010A",company:"STEC",position:"Pantry Assistant",department:"Admin",dateJoined:"2026-03-19",gender:"F",dob:"1956-01-08",nationality:"Singaporean",workPass:"SC",epSpNo:"",workPassIssueDate:"",workPassExpiryDate:"",passportNumber:"",passportIssueDate:"",passportExpiryDate:"",probationEndDate:"2026-09-19",qualification:"PSLE",mobile:"8803 1570",residenceAddress:"Blk 320 Jurong East St #07-72, Singapore 600320",personalEmail:"chongahmui@gmail.com",workEmail:"ahmui@stec.com.sg",bankName:"POSB",bankAccount:"151-0-740300",emergencyContactName:"NG Ah Kow",emergencyContactPhone:"9684 0870",emergencyContactRelation:"Son",phoneAllowance:0,housingAllowance:0,airfareAllowance:0,otherAllowance:0,basicSalary:1900,allowance:0,annualLeave:7,status:"Active",site:"HQ",lastWorkingDay:""},
  {id:"C012",name:"HASSAN Bin Ahmad",fin:"S8268790Z",company:"STEC",position:"Quantity Surveyor",department:"Commercial",dateJoined:"2026-05-01",gender:"M",dob:"1982-06-10",nationality:"Malaysian",workPass:"SPR",epSpNo:"",workPassIssueDate:"",workPassExpiryDate:"",passportNumber:"A1234567",passportIssueDate:"2020-01-01",passportExpiryDate:"2030-01-01",probationEndDate:"2026-11-01",qualification:"Bachelor of Quantity Surveying, UTM",mobile:"9027 9197",residenceAddress:"357A Admiralty Drive, Singapore 751357",personalEmail:"hassanahmad@gmail.com",workEmail:"hassan@stec.com.sg",bankName:"Maybank",bankAccount:"5040-1234-5678",emergencyContactName:"SITI Norbaya",emergencyContactPhone:"8533 9813",emergencyContactRelation:"Spouse",phoneAllowance:50,housingAllowance:0,airfareAllowance:0,otherAllowance:0,basicSalary:5000,allowance:50,annualLeave:14,status:"Active",site:"Jurong Site",lastWorkingDay:""},
];
;

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

const INIT_PROJECTS = [
  {id:"P001",name:"Tuas Industrial Project",code:"TIS-2026",client:"JTC Corporation",startDate:"2026-01-01",endDate:"2027-12-31",status:"Active",supervisorId:"C006",hodId:"C001",members:["C004","C005"],description:"Industrial building construction at Tuas West"},
];

const INIT_COMPANY = {name:"CUSTERA O&M PTE LTD",uen:"202549889D",address:"Singapore"};

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

function Card({children,style={},onClick,glass=false}){
  const [h,sH]=useState(false);
  return<div onClick={onClick} onMouseEnter={()=>onClick&&sH(true)} onMouseLeave={()=>onClick&&sH(false)}
    style={{background:glass?"rgba(255,255,255,0.85)":C.card,backdropFilter:glass?"blur(12px)":"none",borderRadius:16,border:`1px solid ${h?C.accent+"55":C.border}`,padding:"18px 22px",boxShadow:h?"0 8px 32px rgba(99,102,241,0.13)":"0 1px 3px rgba(0,0,0,0.06)",transition:"all 0.18s ease",...(onClick?{cursor:"pointer"}:{}),...style}}>{children}</div>;
}
function PageTitle({title,sub,actions}){
  return<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10,marginBottom:20}}>
    <div><h1 style={{fontSize:20,fontWeight:800,margin:0,color:C.text,letterSpacing:"-0.02em"}}>{title}</h1>{sub&&<p style={{color:C.muted,fontSize:13,margin:"3px 0 0"}}>{sub}</p>}</div>
    {actions&&<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{actions}</div>}
  </div>;
}
function SecTitle({children,icon}){
  return<div style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:12,display:"flex",alignItems:"center",gap:6}}>
    {icon&&<span style={{fontSize:13}}>{icon}</span>}{children}
  </div>;
}

function StatCard({label,value,sub,icon,color,onClick}){
  const [h,sH]=useState(false);
  return<div onClick={onClick} onMouseEnter={()=>onClick&&sH(true)} onMouseLeave={()=>onClick&&sH(false)}
    style={{background:C.card,borderRadius:16,border:`1px solid ${h?color+"55":C.border}`,padding:"18px 20px",boxShadow:h?`0 8px 24px ${color}22`:"0 1px 3px rgba(0,0,0,0.06)",transition:"all 0.18s",flex:1,minWidth:130,cursor:onClick?"pointer":"default"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
      <div><div style={{fontSize:11,fontWeight:600,color:C.muted,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:6}}>{label}</div>
      <div style={{fontSize:26,fontWeight:900,color:C.text,lineHeight:1,letterSpacing:"-0.02em"}}>{value}</div>
      {sub&&<div style={{fontSize:11,color,marginTop:5,fontWeight:600}}>{sub}</div>}</div>
      <div style={{width:46,height:46,borderRadius:13,background:`linear-gradient(135deg,${color}25,${color}10)`,border:`1.5px solid ${color}33`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:19,flexShrink:0}}>{icon}</div>
    </div>
  </div>;
}

function Btn({children,onClick,v="primary",sm=false,style={},disabled=false}){
  const base={border:"none",borderRadius:10,fontWeight:600,cursor:disabled?"not-allowed":"pointer",opacity:disabled?0.5:1,fontSize:sm?11:13,padding:sm?"5px 13px":"9px 18px",fontFamily:"inherit",whiteSpace:"nowrap",transition:"all 0.15s",letterSpacing:"0.01em",...style};
  const V={
    primary:{background:`linear-gradient(135deg,${C.accent} 0%,${C.accentT} 100%)`,color:"#fff",boxShadow:"0 2px 8px rgba(99,102,241,0.28)"},
    danger:{background:`linear-gradient(135deg,${C.danger} 0%,#b91c1c 100%)`,color:"#fff",boxShadow:"0 2px 6px rgba(239,68,68,0.22)"},
    ghost:{background:"rgba(0,0,0,0.03)",color:C.muted,border:`1px solid ${C.border}`},
    success:{background:`linear-gradient(135deg,${C.success} 0%,#047857 100%)`,color:"#fff",boxShadow:"0 2px 6px rgba(16,185,129,0.22)"},
    outline:{background:"transparent",color:C.accent,border:`1.5px solid ${C.accent}`},
    warning:{background:`linear-gradient(135deg,${C.warning} 0%,#d97706 100%)`,color:"#fff",boxShadow:"0 2px 6px rgba(245,158,11,0.22)"},
    teal:{background:`linear-gradient(135deg,${C.teal} 0%,#0d9488 100%)`,color:"#fff",boxShadow:"0 2px 6px rgba(20,184,166,0.22)"},
    purple:{background:`linear-gradient(135deg,${C.purple} 0%,#7c3aed 100%)`,color:"#fff",boxShadow:"0 2px 6px rgba(139,92,246,0.22)"},
  };
  return<button disabled={disabled} onClick={e=>{e.stopPropagation();onClick&&onClick(e);}} style={{...base,...(V[v]||V.primary)}}>{children}</button>;
}

const IS={border:`1px solid ${C.border}`,borderRadius:10,padding:"8px 12px",fontSize:13,outline:"none",background:"#fff",color:C.text,width:"100%",boxSizing:"border-box",fontFamily:"inherit",transition:"border-color 0.15s",};
function Inp({label,value,onChange,type="text",placeholder="",required=false,error="",readOnly=false,rows}){
  const [foc,setFoc]=useState(false);
  const ctrl=rows
    ?<textarea value={value||""} onFocus={()=>setFoc(true)} onBlur={()=>setFoc(false)} onChange={e=>onChange&&onChange(e.target.value)} placeholder={placeholder} rows={rows} style={{...IS,resize:"vertical",border:`1px solid ${error?C.danger:foc?C.accent:C.border}`}}/>
    :<input type={type} value={value||""} onFocus={()=>setFoc(true)} onBlur={()=>setFoc(false)} onChange={e=>onChange&&onChange(e.target.value)} placeholder={placeholder} readOnly={readOnly} style={{...IS,border:`1px solid ${error?C.danger:foc?C.accent:C.border}`,background:readOnly?"#F8FAFC":"#fff"}}/>;
  return<div style={{display:"flex",flexDirection:"column",gap:4}}>{label&&<label style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.06em"}}>{label}{required&&<span style={{color:C.danger}}> *</span>}</label>}{ctrl}{error&&<span style={{fontSize:11,color:C.danger}}>{error}</span>}</div>;
}
function Sel({label,value,onChange,options,required=false}){
  return<div style={{display:"flex",flexDirection:"column",gap:4}}>{label&&<label style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.06em"}}>{label}{required&&<span style={{color:C.danger}}> *</span>}</label>}<select value={value||""} onChange={e=>onChange&&onChange(e.target.value)} style={{...IS}}><option value="">— Select —</option>{options.map(o=>typeof o==="string"?<option key={o} value={o}>{o}</option>:<option key={o.v} value={o.v}>{o.l}</option>)}</select></div>;
}
function Grid({cols=2,children,style={}}){return<div style={{display:"grid",gridTemplateColumns:`repeat(${cols},1fr)`,gap:"12px 16px",...style}}>{children}</div>;}
function FormSection({title,icon,children}){
  return<div style={{marginBottom:20}}>
    <div style={{fontSize:11,fontWeight:700,color:C.accent,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:10,paddingBottom:8,borderBottom:`2px solid ${C.accentL}`,display:"flex",alignItems:"center",gap:6}}>
      {icon&&<span style={{fontSize:14}}>{icon}</span>}{title}
    </div>{children}
  </div>;
}

const STATUS_STYLES={Active:{bg:"#D1FAE5",c:"#059669",dot:"#10B981"},"On Leave":{bg:"#FEF3C7",c:"#D97706",dot:"#F59E0B"},Present:{bg:"#D1FAE5",c:"#059669",dot:"#10B981"},Absent:{bg:"#FEE2E2",c:"#EF4444",dot:"#DC2626"},Pending:{bg:"#FEF3C7",c:"#D97706",dot:"#F59E0B"},Approved:{bg:"#D1FAE5",c:"#059669",dot:"#10B981"},Rejected:{bg:"#FEE2E2",c:"#EF4444",dot:"#DC2626"},"Pending Supervisor":{bg:"#FEF3C7",c:"#B45309",dot:"#F59E0B"},"Pending HR":{bg:"#EDE9FE",c:"#7C3AED",dot:"#8B5CF6"},Published:{bg:"#D1FAE5",c:"#059669",dot:"#10B981"},Draft:{bg:"#F1F5F9",c:"#64748B",dot:"#94A3B8"},Processed:{bg:"#CCFBF1",c:"#14B8A6",dot:"#14B8A6"},Upcoming:{bg:"#EEF2FF",c:"#6366F1",dot:"#6366F1"},Completed:{bg:"#D1FAE5",c:"#059669",dot:"#10B981"},Cancelled:{bg:"#FEE2E2",c:"#EF4444",dot:"#DC2626"},Open:{bg:"#FEF3C7",c:"#D97706",dot:"#F59E0B"},Closed:{bg:"#F1F5F9",c:"#64748B",dot:"#94A3B8"},Inactive:{bg:"#F1F5F9",c:"#64748B",dot:"#94A3B8"},SC:{bg:"#DBEAFE",c:"#1D4ED8",dot:"#3B82F6"},SPR:{bg:"#EDE9FE",c:"#7C3AED",dot:"#8B5CF6"},EP:{bg:"#FEF3C7",c:"#D97706",dot:"#F59E0B"},WP:{bg:"#FCE7F3",c:"#DB2777",dot:"#EC4899"},superadmin:{bg:"#FEE2E2",c:"#EF4444",dot:"#DC2626"},admin:{bg:"#DBEAFE",c:"#1D4ED8",dot:"#3B82F6"},supervisor:{bg:"#EDE9FE",c:"#8B5CF6",dot:"#8B5CF6"},employee:{bg:"#F1F5F9",c:"#64748B",dot:"#94A3B8"},Safety:{bg:"#FEE2E2",c:"#EF4444",dot:"#DC2626"},Holiday:{bg:"#FEF3C7",c:"#D97706",dot:"#F59E0B"},Policy:{bg:"#EDE9FE",c:"#8B5CF6",dot:"#8B5CF6"},General:{bg:"#EEF2FF",c:"#6366F1",dot:"#6366F1"},Compliance:{bg:"#FFEDD5",c:"#EA580C",dot:"#F97316"},Urgent:{bg:"#FEE2E2",c:"#EF4444",dot:"#DC2626"}};
function Badge({s}){const st=STATUS_STYLES[s]||{bg:"#F1F5F9",c:"#64748B",dot:"#94A3B8"};return<span style={{display:"inline-flex",alignItems:"center",gap:4,background:st.bg,color:st.c,borderRadius:99,padding:"3px 10px 3px 8px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}><span style={{width:5,height:5,borderRadius:"50%",background:st.dot,flexShrink:0}}/>{s||"—"}</span>;}

function Modal({title,onClose,children,width=560}){
  return<div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.55)",backdropFilter:"blur(4px)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div style={{background:C.card,borderRadius:20,width:"100%",maxWidth:width,maxHeight:"93vh",overflowY:"auto",boxShadow:"0 32px 80px rgba(0,0,0,0.22)"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"18px 24px 14px",borderBottom:`1px solid ${C.border}`,position:"sticky",top:0,background:C.card,zIndex:1,borderRadius:"20px 20px 0 0"}}>
        <h2 style={{margin:0,fontSize:16,fontWeight:800,color:C.text,letterSpacing:"-0.01em"}}>{title}</h2>
        <button onClick={onClose} style={{background:C.bg,border:"none",borderRadius:8,width:30,height:30,fontSize:16,cursor:"pointer",color:C.muted,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700}}>✕</button>
      </div>
      <div style={{padding:"20px 24px"}}>{children}</div>
    </div>
  </div>;
}

function Toast({msg,type,onDone}){
  useEffect(()=>{const tm=setTimeout(onDone,3200);return()=>clearTimeout(tm);},[]);
  const colors={success:{bg:"#10B981",icon:"✓"},error:{bg:"#EF4444",icon:"✕"},info:{bg:"#6366F1",icon:"ℹ"}};
  const t=colors[type]||colors.success;
  return<div style={{position:"fixed",bottom:28,right:28,zIndex:2000,background:t.bg,color:"#fff",borderRadius:14,padding:"12px 20px",fontWeight:600,fontSize:13,boxShadow:"0 12px 32px rgba(0,0,0,0.20)",display:"flex",gap:10,alignItems:"center",animation:"none"}}>
    <span style={{width:22,height:22,borderRadius:"50%",background:"rgba(255,255,255,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12}}>{t.icon}</span>{msg}
  </div>;
}

function Confirm({msg,onOk,onCancel}){return<Modal title="⚠️ Confirm Action" onClose={onCancel} width={380}><p style={{margin:"0 0 20px",color:C.muted,fontSize:14,lineHeight:1.7}}>{msg}</p><div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={onCancel}>Cancel</Btn><Btn v="danger" onClick={onOk}>Yes, Confirm</Btn></div></Modal>;}

function TH({cols}){return<tr style={{background:"linear-gradient(to right,#F8FAFC,#F1F5F9)"}}>{cols.map(c=><th key={c} style={{padding:"10px 14px",textAlign:"left",fontWeight:700,color:C.muted,fontSize:10,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:"0.07em"}}>{c}</th>)}</tr>;}
function TR({children,onClick}){const[h,sH]=useState(false);return<tr style={{background:h?"#FAFAFF":"#fff",transition:"background 0.1s",cursor:onClick?"pointer":"default"}} onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)} onClick={onClick}>{children}</tr>;}
function TD({children,style={}}){return<td style={{padding:"11px 14px",borderBottom:`1px solid ${C.border}`,...style}}>{children}</td>;}
function TTable({cols,rows,empty="No records yet."}){return<Card style={{padding:0,overflow:"hidden"}}><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}><thead><TH cols={cols}/></thead><tbody>{rows.length===0?<tr><td colSpan={cols.length} style={{padding:40,textAlign:"center",color:C.muted,fontSize:13}}><div style={{fontSize:32,marginBottom:8}}>📭</div>{empty}</td></tr>:rows}</tbody></table></div></Card>;}

function TabBar({tabs,active,onChange}){
  return<div style={{display:"flex",gap:2,borderBottom:`1px solid ${C.border}`,marginBottom:18,background:C.bg,borderRadius:"10px 10px 0 0",padding:"4px 4px 0"}}>
    {tabs.map(tab=>{const iA=active===tab.id;return<button key={tab.id} onClick={()=>onChange(tab.id)} style={{padding:"8px 16px",fontWeight:iA?700:500,fontSize:13,border:"none",background:iA?"#fff":"transparent",cursor:"pointer",color:iA?C.accent:C.muted,borderBottom:iA?`2px solid ${C.accent}`:"2px solid transparent",marginBottom:-1,fontFamily:"inherit",borderRadius:iA?"8px 8px 0 0":"8px 8px 0 0",transition:"all 0.15s",whiteSpace:"nowrap"}}>
      {tab.icon&&<span style={{marginRight:5}}>{tab.icon}</span>}{tab.label}
    </button>;})}
  </div>;
}
function FilterBar({children}){return<div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12,alignItems:"flex-end"}}>{children}</div>;}
function SearchInp({value,onChange,placeholder}){return<div style={{position:"relative",flex:1,minWidth:180}}><span style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:C.muted,fontSize:14}}>🔍</span><input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder||"Search..."} style={{...IS,paddingLeft:34}}/></div>;}
function StatusPill({s,active,onClick}){const st=STATUS_STYLES[s]||{bg:C.bg,c:C.muted};return<button onClick={()=>onClick(s)} style={{border:`1.5px solid ${active?st.c+"66":C.border}`,borderRadius:99,padding:"5px 14px",fontSize:12,fontWeight:active?700:500,background:active?st.bg:"#fff",color:active?st.c:C.muted,cursor:"pointer",fontFamily:"inherit",transition:"all 0.15s"}}>{s}</button>;}

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
  const [form,setForm]=useState({type:leaveTypes[0]?.name||"Annual Leave",from:todayStr(),to:todayStr(),reason:"",halfDay:false,halfDayPeriod:"AM",attachment:""});
  useEffect(()=>{if(leaveTypes.length&&!form.type)setForm(p=>({...p,type:leaveTypes[0].name}));},[leaveTypes]);
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const myLeaves=leaves.filter(l=>l.empId===empId).sort((a,b)=>b.submittedDate.localeCompare(a.submittedDate));
  const approved=myLeaves.filter(l=>l.status==="Approved");
  const alUsed=approved.filter(l=>l.type==="Annual Leave").reduce((s,l)=>s+l.days,0);
  const mlUsed=approved.filter(l=>l.type==="Medical Leave").reduce((s,l)=>s+l.days,0);
  const alEnt=leaveTypes.find(x=>x.name==="Annual Leave")?.days||14;
  const mlEnt=leaveTypes.find(x=>x.name==="Medical Leave")?.days||14;
  function calcDays(f,to,halfDay){const ms=new Date(to)-new Date(f);const d=Math.max(1,Math.floor(ms/86400000)+1);return halfDay?0.5:d;}
  function submit(){
    if(!form.reason.trim())return;
    onAdd({...form,empId,days:calcDays(form.from,form.to,form.halfDay),status:"Pending Supervisor",supervisorComment:"",hrComment:"",submittedDate:todayStr()});
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
      <label style={{display:"flex",gap:8,alignItems:"center",fontSize:13,margin:"10px 0 4px",cursor:"pointer"}}>
        <input type="checkbox" checked={form.halfDay||false} onChange={e=>setForm(p=>({...p,halfDay:e.target.checked}))}/>
        Half Day Leave
      </label>
      {form.halfDay&&<div style={{display:"flex",gap:8,marginBottom:8}}>
        <label style={{display:"flex",gap:6,alignItems:"center",fontSize:13,cursor:"pointer"}}><input type="radio" checked={form.halfDayPeriod==="AM"} onChange={()=>setForm(p=>({...p,halfDayPeriod:"AM"}))}/>AM (Morning)</label>
        <label style={{display:"flex",gap:6,alignItems:"center",fontSize:13,cursor:"pointer"}}><input type="radio" checked={form.halfDayPeriod==="PM"} onChange={()=>setForm(p=>({...p,halfDayPeriod:"PM"}))}/>PM (Afternoon)</label>
      </div>}
      <div style={{background:C.accentL,borderRadius:7,padding:"8px 12px",fontSize:13,color:C.accentT,fontWeight:600,margin:"4px 0 10px"}}>
        {t(lang,"days")}: {form.halfDay?"0.5 (Half Day)":calcDays(form.from,form.to)}
      </div>
      <Inp label={t(lang,"reason")} value={form.reason} onChange={v=>set("reason",v)} rows={3}/>
      {(form.type==="Medical Leave"||form.type==="Hospitalisation Leave")&&<div style={{marginTop:10}}>
        <label style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:6}}>📎 Medical Certificate (required)</label>
        <Btn v="ghost" sm onClick={()=>{const inp=document.createElement("input");inp.type="file";inp.accept="image/*,application/pdf";inp.onchange=e=>{const f=e.target.files?.[0];if(!f)return;const r=new FileReader();r.onload=ev=>set("attachment",ev.target.result);r.readAsDataURL(f);};inp.click();}}>Attach MC / Receipt</Btn>
        {form.attachment&&<span style={{fontSize:11,color:C.success,marginLeft:8}}>✓ File attached</span>}
      </div>}
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
      <div style={{marginTop:12}}><Btn v="teal" onClick={()=>openPayslipPDF(emp,p,{name:"CUSTERA O&M PTE LTD",uen:"202549889D"})}>🖨️ Download PDF</Btn></div>
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

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: ATTENDANCE — Add / Edit / Delete
// ═══════════════════════════════════════════════════════════════════════════════
function AdminAttendance({attendance,employees,onAdd,onEdit,onDelete,lang}){
  const [dateF,setDateF]=useState(todayStr());
  const [search,setSearch]=useState("");
  const [showForm,setShowForm]=useState(false);
  const [editRec,setEditRec]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const EF={empId:"",date:todayStr(),clockIn:"07:00",clockOut:"17:00",site:"",status:"Present",lat:"",lng:""};
  const [f,setF]=useState(EF);
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  const recs=attendance.filter(a=>(!dateF||a.date===dateF)&&(!search||(employees.find(e=>e.id===a.empId)?.name||"").toLowerCase().includes(search.toLowerCase())));
  function calcHrs(ci,co){if(!ci||!co)return 0;const[h1,m1]=ci.split(":").map(Number);const[h2,m2]=co.split(":").map(Number);return Math.round(((h2*60+m2)-(h1*60+m1))/60*100)/100;}
  function openEdit(r){setEditRec(r);setF({empId:r.empId,date:r.date,clockIn:r.clockIn||"",clockOut:r.clockOut||"",site:r.site||"",status:r.status||"Present",lat:r.lat||"",lng:r.lng||""});setShowForm(true);}
  function openNew(){setEditRec(null);setF(EF);setShowForm(true);}
  function save(){const hrs=calcHrs(f.clockIn,f.clockOut);if(editRec){onEdit(editRec.id,{...f,hoursWorked:hrs,id:editRec.id});}else{onAdd({...f,hoursWorked:hrs});}setShowForm(false);}
  function liveHrs(){return Math.max(0,calcHrs(f.clockIn,f.clockOut));}
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"attendance")} actions={[
      <Btn onClick={openNew}>+ Add Record</Btn>,
      <Btn v="ghost" onClick={()=>exportXLS(recs.map(a=>{const e=employees.find(x=>x.id===a.empId)||{};return{Date:fmtDate(a.date),"Emp ID":a.empId,Name:e.name||"",Site:a.site||"","Clock In":a.clockIn,"Clock Out":a.clockOut,"Hours":a.hoursWorked,GPS:`${a.lat||""},${a.lng||""}`,Status:a.status};}),"Attendance",`attendance_${dateF||"all"}.xlsx`)}>{t(lang,"export")}</Btn>
    ]}/>
    <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
      <StatCard label="Present" value={recs.filter(r=>r.status==="Present").length} icon="✓" color={C.success}/>
      <StatCard label="Absent" value={recs.filter(r=>r.status==="Absent").length} icon="✕" color={C.danger}/>
      <StatCard label="Total Records" value={recs.length} icon="R" color={C.accent}/>
    </div>
    <FilterBar>
      <SearchInp value={search} onChange={setSearch} placeholder="Search employee..."/>
      <input type="date" value={dateF} onChange={e=>setDateF(e.target.value)} style={{...IS,width:160}}/>
      <Btn v="ghost" sm onClick={()=>setDateF("")}>All Dates</Btn>
    </FilterBar>
    <TTable cols={["Employee","Date","Site","Clock In","Clock Out","Hours","GPS","Status","Actions"]}
      rows={recs.map(a=>{const emp=employees.find(e=>e.id===a.empId);return<TR key={a.id}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={emp?.name||"?"} size={26}/><span style={{fontWeight:600,fontSize:12}}>{emp?.name||a.empId}</span></div></TD>
        <TD style={{fontSize:12}}>{fmtDate(a.date)}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{a.site||"—"}</TD>
        <TD style={{fontWeight:600,color:C.success}}>{a.clockIn||"—"}</TD>
        <TD style={{fontWeight:600,color:C.danger}}>{a.clockOut||"—"}</TD>
        <TD style={{fontWeight:700,color:C.teal}}>{a.hoursWorked?a.hoursWorked+"h":"—"}</TD>
        <TD style={{fontSize:11,color:C.muted}}>{a.lat?`${Number(a.lat).toFixed(4)},${Number(a.lng).toFixed(4)}`:"—"}</TD>
        <TD><Badge s={a.status}/></TD>
        <TD><div style={{display:"flex",gap:4}}>
          <Btn v="outline" sm onClick={()=>openEdit(a)}>Edit</Btn>
          <Btn v="danger" sm onClick={()=>setConfirmId(a.id)}>Del</Btn>
        </div></TD>
      </TR>;})}/>
    {showForm&&<Modal title={editRec?"Edit Attendance Record":"Add Attendance Record"} onClose={()=>setShowForm(false)} width={520}>
      <Grid>
        <Sel label="Employee" value={f.empId} onChange={v=>setFk("empId",v)} required options={employees.map(e=>({v:e.id,l:`${e.id} - ${e.name}`}))}/>
        <Inp label="Date" type="date" value={f.date} onChange={v=>setFk("date",v)}/>
        <Inp label="Clock In" type="time" value={f.clockIn} onChange={v=>setFk("clockIn",v)}/>
        <Inp label="Clock Out" type="time" value={f.clockOut} onChange={v=>setFk("clockOut",v)}/>
        <Inp label="Work Site" value={f.site} onChange={v=>setFk("site",v)} placeholder="e.g. Tuas Site"/>
        <Sel label="Status" value={f.status} onChange={v=>setFk("status",v)} options={["Present","Absent","Half Day","On Leave"]}/>
        <Inp label="GPS Latitude" value={f.lat} onChange={v=>setFk("lat",v)} placeholder="e.g. 1.3521"/>
        <Inp label="GPS Longitude" value={f.lng} onChange={v=>setFk("lng",v)} placeholder="e.g. 103.8198"/>
      </Grid>
      {f.clockIn&&f.clockOut&&<div style={{background:C.accentL,borderRadius:7,padding:"8px 12px",fontSize:13,color:C.accentT,fontWeight:600,marginTop:10}}>
        Calculated hours worked: {liveHrs()}h
      </div>}
      <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:14}}>
        <Btn v="ghost" onClick={()=>setShowForm(false)}>Cancel</Btn>
        <Btn onClick={save}>Save</Btn>
      </div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete this attendance record?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// APPROVAL CHAIN PROGRESS (3-level visual)
// ═══════════════════════════════════════════════════════════════════════════════
function ApprovalProgress({status,supervisorComment,hrComment}){
  const steps=[
    {label:"Submitted",done:true,color:C.success},
    {label:"Supervisor",done:status!=="Pending Supervisor",active:status==="Pending Supervisor",color:status==="Pending Supervisor"?C.warning:status==="Rejected"&&supervisorComment?C.danger:C.success},
    {label:"HR Admin",done:status==="Approved"||status==="Rejected",active:status==="Pending HR",color:status==="Pending HR"?C.warning:status==="Approved"?C.success:status==="Rejected"&&hrComment?C.danger:C.muted},
    {label:status==="Approved"?"Approved":status==="Rejected"?"Rejected":"Final",done:status==="Approved"||status==="Rejected",color:status==="Approved"?C.success:status==="Rejected"?C.danger:C.muted},
  ];
  return <div style={{display:"flex",alignItems:"center",gap:0}}>
    {steps.map((s,i)=><React.Fragment key={i}>
      <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
        <div style={{width:22,height:22,borderRadius:"50%",background:s.done||s.active?s.color:C.border,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"#fff",flexShrink:0}}>{s.done&&!s.active?"✓":i+1}</div>
        <div style={{fontSize:9,fontWeight:600,color:s.done||s.active?s.color:C.muted,whiteSpace:"nowrap"}}>{s.label}</div>
      </div>
      {i<3&&<div style={{width:20,height:2,background:steps[i+1].done?C.success:C.border,margin:"0 2px",marginBottom:12,flexShrink:0}}/>}
    </React.Fragment>)}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: LEAVE MANAGEMENT — 3-Level Approval
// ═══════════════════════════════════════════════════════════════════════════════
function AdminLeave({leaves,employees,leaveTypes,users,onAction,onAdd,onDelete,onAddType,onEditType,onDeleteType,session,lang}){
  const [tab,setTab]=useState("requests");
  const [filter,setFilter]=useState("All");
  const [detail,setDetail]=useState(null);
  const [rejectId,setRejectId]=useState(null);
  const [rejectComment,setRejectComment]=useState("");
  const [confirmId,setConfirmId]=useState(null);
  const [showApply,setShowApply]=useState(false);
  const [applyForm,setApplyForm]=useState({empId:"",type:leaveTypes[0]?.name||"Annual Leave",from:todayStr(),to:todayStr(),reason:""});
  const isSup=session.role==="supervisor";
  const isHR=session.role==="admin"||session.role==="superadmin";
  const myTeamIds=isSup?users.filter(u=>u.reportingOfficerId===session.id).map(u=>u.empId):null;
  const visible=isSup?leaves.filter(l=>myTeamIds.includes(l.empId)):leaves;
  const filtered=filter==="All"?visible:visible.filter(l=>l.status===filter);

  const pendingSup=visible.filter(l=>l.status==="Pending Supervisor").length;
  const pendingHR=visible.filter(l=>l.status==="Pending HR").length;

  function canAct(l){
    if(isSup&&l.status==="Pending Supervisor")return true;
    if(isHR&&l.status==="Pending HR")return true;
    if(session.role==="superadmin")return l.status==="Pending Supervisor"||l.status==="Pending HR";
    return false;
  }
  function approve(l){
    if(isSup&&l.status==="Pending Supervisor")onAction(l.id,"Pending HR",session.name,"");
    else if(isHR&&l.status==="Pending HR")onAction(l.id,"Approved",session.name,"");
    else if(session.role==="superadmin"){
      if(l.status==="Pending Supervisor")onAction(l.id,"Pending HR",session.name,"");
      else onAction(l.id,"Approved",session.name,"");
    }
    setDetail(null);
  }
  function submitReject(){onAction(rejectId,"Rejected",session.name,rejectComment);setRejectId(null);setRejectComment("");setDetail(null);}
  function calcDays(f,to){return Math.max(1,Math.floor((new Date(to)-new Date(f))/86400000)+1);}
  function submitApply(){if(!applyForm.empId||!applyForm.reason.trim())return;onAdd({...applyForm,days:calcDays(applyForm.from,applyForm.to),status:"Pending Supervisor",supervisorComment:"",hrComment:"",submittedDate:todayStr()});setShowApply(false);}

  const STATUS_FILTERS=["All","Pending Supervisor","Pending HR","Approved","Rejected"];
  const TABS=[{id:"requests",label:"Leave Requests"},{id:"types",label:"Leave Types"}];
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"leaves")} actions={[
      <Btn onClick={()=>setShowApply(true)}>+ Apply Leave</Btn>,
      <Btn v="ghost" onClick={()=>exportXLS(filtered.map(l=>{const e=employees.find(x=>x.id===l.empId)||{};return{"Emp ID":l.empId,Name:e.name||"",Type:l.type,From:fmtDate(l.from),To:fmtDate(l.to),Days:l.days,Reason:l.reason,Status:l.status,"Sup Comment":l.supervisorComment||"","HR Comment":l.hrComment||""};}), "Leaves",`leaves_${todayStr()}.xlsx`)}>{t(lang,"export")}</Btn>
    ]}/>

    {/* 3-Level Approval Summary */}
    <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
      <StatCard label="Pending Supervisor" value={pendingSup} icon="1" color={C.warning} sub="Step 1 of 3"/>
      <StatCard label="Pending HR" value={pendingHR} icon="2" color={C.purple} sub="Step 2 of 3"/>
      <StatCard label="Approved" value={visible.filter(l=>l.status==="Approved").length} icon="✓" color={C.success}/>
      <StatCard label="Rejected" value={visible.filter(l=>l.status==="Rejected").length} icon="✕" color={C.danger}/>
    </div>

    {/* Approval role indicator */}
    <div style={{background:isSup?C.warningL:C.accentL,borderRadius:8,padding:"10px 14px",fontSize:13,fontWeight:600,color:isSup?C.warningD:C.accentT}}>
      {isSup?"Your role: Supervisor — You approve Step 1 (Pending Supervisor)":"Your role: HR Admin — You approve Step 2 (Pending HR)"}
    </div>

    <TabBar tabs={TABS} active={tab} onChange={setTab}/>

    {tab==="requests"&&<>
      <FilterBar>{STATUS_FILTERS.map(s=><StatusPill key={s} s={s} active={filter===s} onClick={setFilter}/>)}</FilterBar>
      <TTable cols={["Employee","Type","From","To","Days","Approval Progress","Status","Actions"]}
        rows={filtered.sort((a,b)=>{const ord={"Pending Supervisor":0,"Pending HR":1,"Approved":2,"Rejected":3};return(ord[a.status]||4)-(ord[b.status]||4);}).map(l=>{const emp=employees.find(e=>e.id===l.empId);return<TR key={l.id} onClick={()=>setDetail(l)}>
          <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={emp?.name||"?"} size={26}/><div><div style={{fontWeight:600,fontSize:12}}>{emp?.name||l.empId}</div><div style={{fontSize:10,color:C.muted}}>{fmtDate(l.submittedDate)}</div></div></div></TD>
          <TD style={{fontSize:12}}>{l.type}</TD>
          <TD style={{fontSize:12,whiteSpace:"nowrap"}}>{fmtDate(l.from)}</TD>
          <TD style={{fontSize:12,whiteSpace:"nowrap"}}>{fmtDate(l.to)}</TD>
          <TD style={{fontWeight:700,textAlign:"center"}}>{l.days}</TD>
          <TD><ApprovalProgress status={l.status} supervisorComment={l.supervisorComment} hrComment={l.hrComment}/></TD>
          <TD><Badge s={l.status}/></TD>
          <TD><div style={{display:"flex",gap:4}} onClick={e=>e.stopPropagation()}>
            {canAct(l)&&<><Btn v="success" sm onClick={()=>approve(l)}>Approve</Btn><Btn v="danger" sm onClick={()=>{setRejectId(l.id);setRejectComment("");}}>Reject</Btn></>}
            <Btn v="ghost" sm onClick={()=>setConfirmId(l.id)}>Del</Btn>
          </div></TD>
        </TR>;})}/>
    </>}
    {tab==="types"&&<LeaveTypeAdmin types={leaveTypes} onAdd={onAddType} onEdit={onEditType} onDelete={onDeleteType} lang={lang}/>}

    {/* Detail modal */}
    {detail&&<Modal title="Leave Request Detail" onClose={()=>setDetail(null)} width={520}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
        {[["Employee",employees.find(e=>e.id===detail.empId)?.name||detail.empId],["Leave Type",detail.type],["From",fmtDate(detail.from)],["To",fmtDate(detail.to)],["Days",detail.days],["Submitted",fmtDate(detail.submittedDate)]].map(([k,v])=><div key={k} style={{background:C.bg,borderRadius:7,padding:"8px 12px"}}><div style={{fontSize:11,color:C.muted}}>{k}</div><div style={{fontWeight:700}}>{v}</div></div>)}
      </div>
      <div style={{background:C.bg,borderRadius:7,padding:"10px 12px",marginBottom:12}}><div style={{fontSize:11,color:C.muted,marginBottom:4}}>Reason</div><div style={{fontSize:13}}>{detail.reason}</div></div>
      <div style={{marginBottom:14}}><div style={{fontSize:12,fontWeight:700,color:C.muted,marginBottom:8}}>3-Level Approval Progress</div><ApprovalProgress status={detail.status} supervisorComment={detail.supervisorComment} hrComment={detail.hrComment}/></div>
      {detail.supervisorComment&&<div style={{background:C.warningL,borderRadius:7,padding:"8px 12px",marginBottom:8,fontSize:12}}><strong>Supervisor comment:</strong> {detail.supervisorComment}</div>}
      {detail.hrComment&&<div style={{background:C.accentL,borderRadius:7,padding:"8px 12px",marginBottom:8,fontSize:12}}><strong>HR comment:</strong> {detail.hrComment}</div>}
      {canAct(detail)&&<div style={{display:"flex",gap:8,marginTop:12}}><Btn v="success" onClick={()=>approve(detail)}>Approve</Btn><Btn v="danger" onClick={()=>{setRejectId(detail.id);setRejectComment("");}}>Reject</Btn></div>}
    </Modal>}

    {/* Reject with comment */}
    {rejectId&&<Modal title="Reject Leave — Add Comment" onClose={()=>setRejectId(null)} width={420}>
      <Inp label="Reason for rejection" value={rejectComment} onChange={setRejectComment} rows={3} placeholder="Please provide a reason..."/>
      <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:12}}><Btn v="ghost" onClick={()=>setRejectId(null)}>Cancel</Btn><Btn v="danger" onClick={submitReject}>Confirm Reject</Btn></div>
    </Modal>}

    {/* Apply leave */}
    {showApply&&<Modal title="Apply Leave on Behalf" onClose={()=>setShowApply(false)} width={480}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Sel label="Employee" value={applyForm.empId} onChange={v=>setApplyForm(p=>({...p,empId:v}))} required options={employees.map(e=>({v:e.id,l:`${e.id} - ${e.name}`}))}/>
        <Sel label="Leave Type" value={applyForm.type} onChange={v=>setApplyForm(p=>({...p,type:v}))} options={leaveTypes.map(l=>l.name)}/>
        <Grid><Inp label="From" type="date" value={applyForm.from} onChange={v=>setApplyForm(p=>({...p,from:v}))}/><Inp label="To" type="date" value={applyForm.to} onChange={v=>setApplyForm(p=>({...p,to:v}))}/></Grid>
        <Inp label="Reason" value={applyForm.reason} onChange={v=>setApplyForm(p=>({...p,reason:v}))} rows={2}/>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowApply(false)}>Cancel</Btn><Btn onClick={submitApply}>Submit</Btn></div>
      </div>
    </Modal>}

    {confirmId&&<Confirm msg="Delete this leave record?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: CLAIMS — 3-Level Approval
// ═══════════════════════════════════════════════════════════════════════════════
function AdminClaims({claims,employees,users,session,onAction,onDelete,lang}){
  const [filter,setFilter]=useState("All");
  const [detail,setDetail]=useState(null);
  const [rejectId,setRejectId]=useState(null);
  const [rejectComment,setRejectComment]=useState("");
  const [confirmId,setConfirmId]=useState(null);
  const isSup=session.role==="supervisor";
  const isHR=session.role==="admin"||session.role==="superadmin";
  const myTeamIds=isSup?users.filter(u=>u.reportingOfficerId===session.id).map(u=>u.empId):null;
  const visible=isSup?claims.filter(c=>myTeamIds.includes(c.empId)):claims;
  const filtered=filter==="All"?visible:visible.filter(c=>c.status===filter);
  const totalApproved=visible.filter(c=>c.status==="Approved").reduce((s,c)=>s+c.amount,0);
  function canAct(c){
    if(isSup&&c.status==="Pending Supervisor")return true;
    if(isHR&&c.status==="Pending HR")return true;
    if(session.role==="superadmin")return c.status==="Pending Supervisor"||c.status==="Pending HR";
    return false;
  }
  function approve(c){
    if(isSup&&c.status==="Pending Supervisor")onAction(c.id,"Pending HR",session.name,"");
    else if(isHR&&c.status==="Pending HR")onAction(c.id,"Approved",session.name,"");
    else if(session.role==="superadmin"){if(c.status==="Pending Supervisor")onAction(c.id,"Pending HR",session.name,"");else onAction(c.id,"Approved",session.name,"");}
    setDetail(null);
  }
  function submitReject(){onAction(rejectId,"Rejected",session.name,rejectComment);setRejectId(null);setRejectComment("");setDetail(null);}
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title={t(lang,"claims")} sub={`Total Approved: ${sgd(totalApproved)}`} actions={[<Btn v="ghost" onClick={()=>exportXLS(filtered.map(c=>{const e=employees.find(x=>x.id===c.empId)||{};return{Name:e.name||"",Type:c.type,Date:fmtDate(c.date),"Amount (S$)":c.amount,Description:c.description,Status:c.status,"Sup Comment":c.supervisorComment||"","HR Comment":c.hrComment||""};}), "Claims",`claims_${todayStr()}.xlsx`)}>{t(lang,"export")}</Btn>]}/>

    <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
      <StatCard label="Pending Supervisor" value={visible.filter(c=>c.status==="Pending Supervisor").length} icon="1" color={C.warning} sub="Step 1 of 3"/>
      <StatCard label="Pending HR" value={visible.filter(c=>c.status==="Pending HR").length} icon="2" color={C.purple} sub="Step 2 of 3"/>
      <StatCard label="Approved" value={visible.filter(c=>c.status==="Approved").length} icon="✓" color={C.success}/>
      <StatCard label="Total Approved" value={sgd(totalApproved)} icon="$" color={C.teal}/>
    </div>

    <div style={{background:isSup?C.warningL:C.accentL,borderRadius:8,padding:"10px 14px",fontSize:13,fontWeight:600,color:isSup?C.warningD:C.accentT}}>
      {isSup?"Your role: Supervisor — You verify Step 1 (Pending Supervisor)":"Your role: HR Admin — You approve Step 2 (Pending HR)"}
    </div>

    <FilterBar>{["All","Pending Supervisor","Pending HR","Approved","Rejected"].map(s=><StatusPill key={s} s={s} active={filter===s} onClick={setFilter}/>)}</FilterBar>

    <TTable cols={["Employee","Type","Date","Amount","Description","Approval Progress","Status","Actions"]}
      rows={filtered.sort((a,b)=>{const ord={"Pending Supervisor":0,"Pending HR":1,"Approved":2,"Rejected":3};return(ord[a.status]||4)-(ord[b.status]||4);}).map(c=>{const emp=employees.find(e=>e.id===c.empId);return<TR key={c.id} onClick={()=>setDetail(c)}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={emp?.name||"?"} size={26}/><span style={{fontWeight:600,fontSize:12}}>{emp?.name||c.empId}</span></div></TD>
        <TD><Badge s={c.type}/></TD>
        <TD style={{fontSize:12}}>{fmtDate(c.date)}</TD>
        <TD style={{fontWeight:700,color:C.success}}>{sgd(c.amount)}</TD>
        <TD style={{fontSize:12,color:C.muted,maxWidth:130}}><span title={c.description}>{c.description.slice(0,30)}{c.description.length>30?"...":""}</span></TD>
        <TD><ApprovalProgress status={c.status} supervisorComment={c.supervisorComment} hrComment={c.hrComment}/></TD>
        <TD><Badge s={c.status}/></TD>
        <TD><div style={{display:"flex",gap:4}} onClick={e=>e.stopPropagation()}>
          {canAct(c)&&<><Btn v="success" sm onClick={()=>approve(c)}>Approve</Btn><Btn v="danger" sm onClick={()=>{setRejectId(c.id);setRejectComment("");}}>Reject</Btn></>}
          <Btn v="ghost" sm onClick={()=>setConfirmId(c.id)}>Del</Btn>
        </div></TD>
      </TR>;})}/>

    {detail&&<Modal title="Claim Detail" onClose={()=>setDetail(null)} width={520}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
        {[["Employee",employees.find(e=>e.id===detail.empId)?.name||detail.empId],["Type",detail.type],["Date",fmtDate(detail.date)],["Amount",sgd(detail.amount)]].map(([k,v])=><div key={k} style={{background:C.bg,borderRadius:7,padding:"8px 12px"}}><div style={{fontSize:11,color:C.muted}}>{k}</div><div style={{fontWeight:700}}>{v}</div></div>)}
      </div>
      <div style={{background:C.bg,borderRadius:7,padding:"10px 12px",marginBottom:12}}><div style={{fontSize:11,color:C.muted,marginBottom:4}}>Description</div><div style={{fontSize:13}}>{detail.description}</div></div>
      <div style={{marginBottom:14}}><div style={{fontSize:12,fontWeight:700,color:C.muted,marginBottom:8}}>3-Level Approval Progress</div><ApprovalProgress status={detail.status} supervisorComment={detail.supervisorComment} hrComment={detail.hrComment}/></div>
      {detail.supervisorComment&&<div style={{background:C.warningL,borderRadius:7,padding:"8px 12px",marginBottom:8,fontSize:12}}><strong>Supervisor comment:</strong> {detail.supervisorComment}</div>}
      {detail.hrComment&&<div style={{background:C.accentL,borderRadius:7,padding:"8px 12px",marginBottom:8,fontSize:12}}><strong>HR comment:</strong> {detail.hrComment}</div>}
      {canAct(detail)&&<div style={{display:"flex",gap:8,marginTop:12}}><Btn v="success" onClick={()=>approve(detail)}>Approve</Btn><Btn v="danger" onClick={()=>{setRejectId(detail.id);setRejectComment("");}}>Reject</Btn></div>}
    </Modal>}
    {rejectId&&<Modal title="Reject Claim — Add Comment" onClose={()=>setRejectId(null)} width={420}>
      <Inp label="Reason for rejection" value={rejectComment} onChange={setRejectComment} rows={3} placeholder="Please provide a reason..."/>
      <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:12}}><Btn v="ghost" onClick={()=>setRejectId(null)}>Cancel</Btn><Btn v="danger" onClick={submitReject}>Confirm Reject</Btn></div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete claim?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

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

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: PAYROLL — Fully Editable with OT / Reimbursement / Claims Breakdown
// ═══════════════════════════════════════════════════════════════════════════════

function PayrollEntryModal({emp,existing,month,claims,onSave,onClose,company}){
  const approvedOT=claims.filter(c=>c.empId===emp.id&&c.type==="OT"&&c.status==="Approved"&&c.date.startsWith(month)).reduce((s,c)=>s+c.amount,0);
  const approvedReimb=claims.filter(c=>c.empId===emp.id&&c.type!=="OT"&&c.status==="Approved"&&c.date.startsWith(month)).reduce((s,c)=>s+c.amount,0);
  const [f,setF]=useState({
    basic:String((existing?.basic??Number(emp.basicSalary))||0),
    phoneAllowance:String((existing?.phoneAllowance??Number(emp.phoneAllowance))||0),
    housingAllowance:String((existing?.housingAllowance??Number(emp.housingAllowance))||0),
    airfareAllowance:String((existing?.airfareAllowance??Number(emp.airfareAllowance))||0),
    otherAllowance:String((existing?.otherAllowance??Number(emp.otherAllowance||emp.allowance))||0),
    otManual:String(existing?.otManual||0),
    commission:String(existing?.commission||0),
    leaveEncashment:String(existing?.leaveEncashment||0),
    nplDeduction:String(existing?.nplDeduction||0),
    incomeTax:String(existing?.incomeTax||0),
    cdac:String(existing?.cdac||0),
    otherDeductions:String((existing?.otherDeductions??0)||0),
    notes:existing?.notes||"",
  });
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  const n=k=>Number(f[k])||0;
  const totalAllowance=n("phoneAllowance")+n("housingAllowance")+n("airfareAllowance")+n("otherAllowance");
  const totalOT=approvedOT+n("otManual");
  const gross=n("basic")+totalAllowance+totalOT+approvedReimb+n("commission")+n("leaveEncashment");
  const cpf=calcCPF(n("basic"),emp.nationality);
  const sdl=calcSDF(n("basic"));
  const totalDeductions=cpf.employee+n("nplDeduction")+n("incomeTax")+n("cdac")+n("otherDeductions");
  const netPay=gross-totalDeductions;
  function syncFromEmployee(){setF(p=>({...p,basic:String(Number(emp.basicSalary)||0),phoneAllowance:String(Number(emp.phoneAllowance)||0),housingAllowance:String(Number(emp.housingAllowance)||0),airfareAllowance:String(Number(emp.airfareAllowance)||0),otherAllowance:String(Number(emp.otherAllowance||emp.allowance)||0)}));} 
  function save(){
    const data={id:existing?.id||uid(),empId:emp.id,month,basic:n("basic"),phoneAllowance:n("phoneAllowance"),housingAllowance:n("housingAllowance"),airfareAllowance:n("airfareAllowance"),otherAllowance:n("otherAllowance"),allowance:totalAllowance,otPay:totalOT,otManual:n("otManual"),otClaims:approvedOT,reimbursement:approvedReimb,commission:n("commission"),leaveEncashment:n("leaveEncashment"),nplDeduction:n("nplDeduction"),incomeTax:n("incomeTax"),cdac:n("cdac"),otherDeductions:n("otherDeductions"),gross,cpfEmployee:cpf.employee,cpfEmployer:cpf.employer,sdl,netPay,status:existing?.status||"Draft",processedOn:todayStr(),publishedOn:existing?.publishedOn||"",notes:f.notes};
    onSave(data);onClose();
  }
  const row=(l,v,c,bold)=><div style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid "+C.border,fontWeight:bold?800:500}}><span style={{color:C.muted,fontSize:12}}>{l}</span><span style={{color:c||C.text,fontSize:bold?14:12}}>{v}</span></div>;
  return <Modal title={`Payroll — ${emp.name} (${month})`} onClose={onClose} width={700}>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <div style={{display:"flex",flexDirection:"column",gap:0}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><div style={{fontSize:11,fontWeight:700,color:C.accent,textTransform:"uppercase"}}>Earnings</div><Btn v="outline" sm onClick={syncFromEmployee}>↺ Sync from Employee</Btn></div>
        <div style={{display:"flex",flexDirection:"column",gap:9}}>
          <Inp label="Basic Salary (S$)" type="number" value={f.basic} onChange={v=>setFk("basic",v)}/>
          <Inp label="Phone Allowance (S$)" type="number" value={f.phoneAllowance} onChange={v=>setFk("phoneAllowance",v)}/>
          <Inp label="Housing Allowance (S$)" type="number" value={f.housingAllowance} onChange={v=>setFk("housingAllowance",v)}/>
          <Inp label="Airfare Allowance (S$)" type="number" value={f.airfareAllowance} onChange={v=>setFk("airfareAllowance",v)}/>
          <Inp label="Other Allowance (S$)" type="number" value={f.otherAllowance} onChange={v=>setFk("otherAllowance",v)}/>
          <div style={{background:C.tealL,borderRadius:7,padding:"8px 12px"}}>
            <div style={{fontSize:10,color:C.teal,fontWeight:700}}>OT from Approved Claims (auto)</div>
            <div style={{fontWeight:800,color:C.teal}}>{sgd(approvedOT)}</div>
          </div>
          <Inp label="Additional OT / Manual OT (S$)" type="number" value={f.otManual} onChange={v=>setFk("otManual",v)}/>
          {approvedReimb>0&&<div style={{background:C.tealL,borderRadius:7,padding:"8px 12px"}}><div style={{fontSize:10,color:C.teal,fontWeight:700}}>Reimbursements from Claims (auto)</div><div style={{fontWeight:800,color:C.teal}}>{sgd(approvedReimb)}</div></div>}
          <Inp label="Commission / Bonus (S$)" type="number" value={f.commission} onChange={v=>setFk("commission",v)}/>
          <Inp label="Leave Encashment (S$)" type="number" value={f.leaveEncashment} onChange={v=>setFk("leaveEncashment",v)}/>
        </div>
        <div style={{fontSize:11,fontWeight:700,color:C.danger,textTransform:"uppercase",margin:"12px 0 8px"}}>Deductions</div>
        <div style={{display:"flex",flexDirection:"column",gap:9}}>
          <div style={{background:C.dangerL,borderRadius:7,padding:"8px 12px"}}><div style={{fontSize:10,color:C.dangerD,fontWeight:700}}>CPF Employee 20% (auto-calculated)</div><div style={{fontWeight:800,color:C.danger}}>{sgd(cpf.employee)}</div></div>
          <Inp label="No Paid Leave Deduction (S$)" type="number" value={f.nplDeduction} onChange={v=>setFk("nplDeduction",v)}/>
          <Inp label="Income Tax / IR8A (S$)" type="number" value={f.incomeTax} onChange={v=>setFk("incomeTax",v)}/>
          <Inp label="CDAC / SINDA / MBMF (S$)" type="number" value={f.cdac} onChange={v=>setFk("cdac",v)}/>
          <Inp label="Other Deductions (S$)" type="number" value={f.otherDeductions} onChange={v=>setFk("otherDeductions",v)}/>
        </div>
        <div style={{marginTop:10}}><Inp label="Notes" value={f.notes} onChange={v=>setFk("notes",v)} rows={2}/></div>
      </div>
      <div>
        <div style={{fontSize:11,fontWeight:700,color:C.accent,textTransform:"uppercase",marginBottom:8}}>Live Payslip Preview</div>
        <div style={{background:C.sidebar,color:"#fff",borderRadius:10,padding:"12px 14px",marginBottom:10}}>
          <div style={{fontWeight:800,fontSize:13}}>{emp.name}</div>
          <div style={{fontSize:11,opacity:0.7}}>{emp.position} · {emp.id}</div>
          <div style={{fontSize:11,opacity:0.7,marginTop:2}}>{month} · {emp.nationality}</div>
        </div>
        {row("Basic Salary",sgd(n("basic")))}
        {row("Phone Allowance",sgd(n("phoneAllowance")))}
        {row("Housing Allowance",sgd(n("housingAllowance")))}
        {row("Airfare Allowance",sgd(n("airfareAllowance")))}
        {n("otherAllowance")>0&&row("Other Allowance",sgd(n("otherAllowance")))}
        {totalOT>0&&row("OT Pay",sgd(totalOT),C.teal)}
        {approvedReimb>0&&row("Reimbursements",sgd(approvedReimb),C.teal)}
        {n("commission")>0&&row("Commission/Bonus",sgd(n("commission")))}
        {n("leaveEncashment")>0&&row("Leave Encashment",sgd(n("leaveEncashment")))}
        {row("GROSS PAY",sgd(gross),C.text,true)}
        <div style={{height:6}}/>
        {row("CPF (Employee 20%)","-"+sgd(cpf.employee),C.danger)}
        {n("nplDeduction")>0&&row("No Paid Leave","-"+sgd(n("nplDeduction")),C.danger)}
        {n("incomeTax")>0&&row("Income Tax","-"+sgd(n("incomeTax")),C.danger)}
        {n("cdac")>0&&row("CDAC/SINDA","-"+sgd(n("cdac")),C.danger)}
        {n("otherDeductions")>0&&row("Other Deductions","-"+sgd(n("otherDeductions")),C.danger)}
        <div style={{background:C.successL,borderRadius:10,padding:"14px",margin:"10px 0",textAlign:"center"}}>
          <div style={{fontSize:11,color:C.successD,fontWeight:700}}>NET PAY</div>
          <div style={{fontSize:26,fontWeight:900,color:C.successD}}>{sgd(netPay)}</div>
        </div>
        <div style={{background:C.bg,borderRadius:8,padding:"10px 12px"}}>
          <div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:6,textTransform:"uppercase"}}>Employer Contributions (not deducted)</div>
          {row("CPF Employer (17%)",sgd(cpf.employer),C.warning)}
          {row("Skills Development Levy",sgd(sdl))}
        </div>
        <div style={{marginTop:10,fontSize:11,color:C.muted}}>CPF applies: {(emp.nationality==="Singaporean"||emp.nationality==="SPR")?"Yes (SC/SPR)":"No — "+emp.workPass}</div>
      </div>
    </div>
    <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:16,paddingTop:14,borderTop:"1px solid "+C.border}}>
      <Btn v="ghost" onClick={onClose}>Cancel</Btn>
      <Btn onClick={save}>{existing?"Update Entry":"Save Entry"}</Btn>
    </div>
  </Modal>;
}

function AdminPayroll({employees,payroll,claims,onProcess,onDelete,onPublish,company,lang}){
  const [month,setMonth]=useState(new Date().toISOString().slice(0,7));
  const [editEntry,setEditEntry]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const [showSlip,setShowSlip]=useState(null);
  const processed=payroll.filter(p=>p.month===month);
  const processedIds=new Set(processed.map(p=>p.empId));
  const activeEmps=employees.filter(e=>e.status==="Active");
  const totGross=processed.reduce((s,p)=>s+p.gross,0);
  const totNet=processed.reduce((s,p)=>s+p.netPay,0);
  const totEmpCPF=processed.reduce((s,p)=>s+p.cpfEmployee,0);
  const totErCPF=processed.reduce((s,p)=>s+p.cpfEmployer,0);
  function processAll(){activeEmps.filter(e=>!processedIds.has(e.id)).forEach(e=>{
    const basic=Number(e.basicSalary)||0;
    const phone=Number(e.phoneAllowance)||0,housing=Number(e.housingAllowance)||0,airfare=Number(e.airfareAllowance)||0,other=Number(e.otherAllowance||e.allowance)||0;
    const totalAllow=phone+housing+airfare+other;
    const otPay=claims.filter(c=>c.empId===e.id&&c.type==="OT"&&c.status==="Approved"&&c.date.startsWith(month)).reduce((s,c)=>s+c.amount,0);
    const reimb=claims.filter(c=>c.empId===e.id&&c.type!=="OT"&&c.status==="Approved"&&c.date.startsWith(month)).reduce((s,c)=>s+c.amount,0);
    const gross=basic+totalAllow+otPay+reimb;const cpf=calcCPF(basic,e.nationality);const sdl=calcSDF(basic);
    onProcess({id:uid(),empId:e.id,month,basic,phoneAllowance:phone,housingAllowance:housing,airfareAllowance:airfare,otherAllowance:other,allowance:totalAllow,otPay,otClaims:otPay,reimbursement:reimb,commission:0,leaveEncashment:0,nplDeduction:0,incomeTax:0,cdac:0,otherDeductions:0,gross,cpfEmployee:cpf.employee,cpfEmployer:cpf.employer,sdl,netPay:gross-cpf.employee,status:"Draft",processedOn:todayStr(),publishedOn:"",notes:""});
  });}
  function doExport(){exportMulti([
    {n:"Payroll",data:processed.map(p=>{const e=employees.find(x=>x.id===p.empId)||{};return{Month:p.month,"Emp ID":p.empId,Name:e.name||"",Company:e.company||"","Work Pass":e.workPass||"","FIN/NRIC":e.fin||"","Basic (S$)":p.basic,"Phone Allow":p.phoneAllowance||0,"Housing Allow":p.housingAllowance||0,"Airfare Allow":p.airfareAllowance||0,"Other Allow":p.otherAllowance||0,"OT (S$)":p.otPay||0,"Reimb (S$)":p.reimbursement||0,"Commission":p.commission||0,"Leave Encash":p.leaveEncashment||0,"Gross (S$)":p.gross,"CPF Ee":p.cpfEmployee,"CPF Er":p.cpfEmployer,"SDL":p.sdl,"NPL Deduct":p.nplDeduction||0,"Income Tax":p.incomeTax||0,"CDAC":p.cdac||0,"Other Deduct":p.otherDeductions||0,"Net Pay (S$)":p.netPay,Status:p.status}; })},
    {n:"CPF Submission",data:processed.filter(p=>p.cpfEmployee>0).map(p=>{const e=employees.find(x=>x.id===p.empId)||{};return{Month:p.month,Name:e.name||"","FIN/NRIC":e.fin||"","Ordinary Wages":p.basic,"Additional Wages":(p.allowance||0)+(p.otPay||0),"Ee CPF":p.cpfEmployee,"Er CPF":p.cpfEmployer,"Total CPF":p.cpfEmployee+p.cpfEmployer};})},
  ],`payroll_${month}.xlsx`);}
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
      <div><h1 style={{fontSize:19,fontWeight:800,margin:0,color:C.text}}>Payroll</h1><input type="month" value={month} onChange={e=>setMonth(e.target.value)} style={{...IS,width:160,marginTop:6}}/></div>
      <div style={{display:"flex",gap:8}}><Btn v="ghost" onClick={doExport}>Export Excel</Btn><Btn v="warning" onClick={processAll} disabled={activeEmps.filter(e=>!processedIds.has(e.id)).length===0}>Process All (Default)</Btn></div>
    </div>
    <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
      <StatCard label="Total Gross" value={sgd(totGross)} icon="$" color={C.accent}/>
      <StatCard label="Total Net Pay" value={sgd(totNet)} icon="N" color={C.success}/>
      <StatCard label="CPF (Employee)" value={sgd(totEmpCPF)} icon="E" color={C.purple}/>
      <StatCard label="CPF (Employer)" value={sgd(totErCPF)} icon="R" color={C.warning}/>
    </div>
    <TTable cols={["Employee","Basic","Allowances","OT","Gross","CPF Ee","Net Pay","Status","Actions"]}
      rows={activeEmps.map(e=>{const p=processed.find(x=>x.empId===e.id);
      const basic=Number(e.basicSalary)||0;
      const totalAllow=(Number(e.phoneAllowance)||0)+(Number(e.housingAllowance)||0)+(Number(e.airfareAllowance)||0)+(Number(e.otherAllowance||e.allowance)||0);
      const cpf=calcCPF(basic,e.nationality);
      return<TR key={e.id}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={e.name} size={26}/><div><div style={{fontWeight:600,fontSize:12}}>{e.name}</div><div style={{fontSize:10,color:C.muted}}>{e.id} · {e.nationality}</div></div></div></TD>
        <TD style={{fontSize:12}}>{sgd(p?.basic||basic)}</TD>
        <TD style={{fontSize:12,color:C.purple}}>{sgd(p?.allowance||totalAllow)}</TD>
        <TD style={{fontSize:12,color:C.teal}}>{sgd(p?.otPay||0)}</TD>
        <TD style={{fontWeight:600}}>{sgd(p?.gross||basic+totalAllow)}</TD>
        <TD style={{color:C.danger,fontSize:12}}>{sgd(p?.cpfEmployee||cpf.employee)}</TD>
        <TD style={{fontWeight:700,color:C.success}}>{sgd(p?.netPay||(basic+totalAllow-cpf.employee))}</TD>
        <TD><Badge s={p?.status||"Draft"}/></TD>
        <TD><div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
          <Btn v={p?"outline":"primary"} sm onClick={()=>setEditEntry({emp:e,existing:p})}>{p?"Edit":"Process"}</Btn>
          {p&&<Btn v="ghost" sm onClick={()=>setShowSlip({emp:e,p})}>Slip</Btn>}
          {p&&p.status==="Draft"&&<Btn v="success" sm onClick={()=>onPublish(p.id)}>Publish</Btn>}
          {p&&p.status==="Published"&&<Btn v="teal" sm onClick={()=>openPayslipPDF(e,p,company||{})}>PDF</Btn>}
          {p&&<Btn v="danger" sm onClick={()=>setConfirmId(p.id)}>Del</Btn>}
        </div></TD>
      </TR>;})}/>
    {editEntry&&<PayrollEntryModal emp={editEntry.emp} existing={editEntry.existing} month={month} claims={claims} company={company||{}}
      onSave={d=>{if(editEntry.existing)onDelete(editEntry.existing.id);onProcess({...d,id:uid()});setEditEntry(null);}}
      onClose={()=>setEditEntry(null)}/>}
    {showSlip&&<Modal title="Payslip Preview" onClose={()=>setShowSlip(null)} width={460}>
      <EmpPayslip empId={showSlip.emp.id} payroll={[{...showSlip.p,status:"Published"}]} employees={employees} lang={lang} company={company}/>
      <div style={{marginTop:12}}><Btn v="teal" onClick={()=>openPayslipPDF(showSlip.emp,showSlip.p,company||{})}>🖨️ Download PDF (Custera Template)</Btn></div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete payroll record?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

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

// ═══════════════════════════════════════════════════════════════════════════════
// ERROR BOUNDARY — prevents blank page crashes, shows friendly error instead
// ═══════════════════════════════════════════════════════════════════════════════
class ErrorBoundary extends React.Component {
  constructor(props){super(props);this.state={err:null};}
  static getDerivedStateFromError(e){return{err:e};}
  componentDidCatch(e,info){console.error("Page error:",e,info);}
  render(){
    if(this.state.err){return(
      <div style={{padding:32,textAlign:"center"}}>
        <div style={{fontSize:40,marginBottom:16}}>⚠️</div>
        <div style={{fontWeight:800,fontSize:16,marginBottom:8,color:C.danger}}>Something went wrong</div>
        <div style={{color:C.muted,fontSize:13,marginBottom:20,maxWidth:400,margin:"0 auto 20px"}}>{this.state.err.message}</div>
        <button onClick={()=>this.setState({err:null})} style={{background:C.accent,color:"#fff",border:"none",borderRadius:8,padding:"8px 20px",cursor:"pointer",fontWeight:600,fontSize:14}}>Try Again</button>
      </div>
    );}
    return this.props.children;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PAYSLIP PDF GENERATOR (Custera Template)
// ═══════════════════════════════════════════════════════════════════════════════
function fmtSGD2(n){return"$"+Number(n||0).toLocaleString("en-SG",{minimumFractionDigits:2,maximumFractionDigits:2});}
function openPayslipPDF(emp,p,co){
  const[yr,mo]=p.month.split("-");const lastDay=new Date(parseInt(yr),parseInt(mo),0).getDate();
  const moName=new Date(parseInt(yr),parseInt(mo)-1,1).toLocaleDateString("en-SG",{month:"short"});
  const totalDeductions=(p.cpfEmployee||0)+(p.incomeTax||0)+(p.nplDeduction||0)+(p.otherDeductions||0)+(p.cdac||0);
  const html=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Payslip-${emp.name}-${p.month}</title>
<style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:Arial,sans-serif;padding:32px;color:#333;font-size:11px;}
.logo{font-size:26px;font-weight:900;color:#1A2340;margin-bottom:16px;letter-spacing:-1px;}
.logo span{color:#E55B1F;}.bar{display:inline-block;width:32px;height:5px;background:linear-gradient(to right,#E55B1F,#F5A623,#2ECC71);vertical-align:middle;margin:0 2px;}
.title{background:#2E86C1;color:#fff;text-align:center;padding:10px;font-size:16px;font-weight:bold;letter-spacing:3px;margin:14px 0 18px;}
.hgrid{display:table;width:100%;margin-bottom:18px;}.hcol{display:table-cell;width:50%;vertical-align:top;}
.hr{margin-bottom:5px;font-size:11px;}.hl{display:inline-block;width:140px;color:#666;}
.wrap{display:table;width:100%;border:1px solid #ccc;}.ec{display:table-cell;width:50%;vertical-align:top;}
.ec:first-child{border-right:1px solid #ccc;}table{width:100%;border-collapse:collapse;}
th{background:#2E86C1;color:#fff;padding:7px 10px;text-align:left;font-size:11px;}th.r{text-align:right;}
td{padding:5px 10px;border-bottom:1px solid #f0f0f0;font-size:11px;}td.r{text-align:right;}
tr.b td{font-weight:bold;border-top:2px solid #bbb;background:#f8f8f8;}
.sum{float:right;width:260px;border:1px solid #ccc;margin-top:10px;}
.sr{display:flex;justify-content:space-between;padding:5px 12px;border-bottom:1px solid #eee;font-size:11px;}
.sr.net{background:#e8f5fd;font-weight:bold;font-size:13px;padding:8px 12px;}
.ft{clear:both;text-align:center;margin-top:50px;font-size:10px;color:#888;border-top:1px solid #ddd;padding-top:12px;}
@media print{.np{display:none;}}
</style></head><body>
<button class="np" onclick="window.print()" style="margin-bottom:16px;padding:8px 20px;background:#2563EB;color:#fff;border:none;cursor:pointer;border-radius:5px;font-size:13px;font-weight:600;">🖨️ Print / Save as PDF</button>
<div class="logo">CUSTER<span>A</span><span class="bar"></span></div>
<div class="title">PAYSLIP</div>
<div class="hgrid"><div class="hcol">
<div class="hr"><span class="hl">Company Name</span>${co?.name||"CUSTERA O&M PTE LTD"}</div>
<div class="hr"><span class="hl">UEN</span>${co?.uen||"202549889D"}</div>
<div class="hr"><span class="hl">Payment Date</span>${lastDay}-${moName}-${yr}</div>
<div class="hr"><span class="hl">Salary Period</span>01-${moName}-${yr} to ${lastDay}-${moName}-${yr}</div>
<div class="hr"><span class="hl">Last Working Day</span>${emp.lastWorkingDay||"-"}</div>
</div><div class="hcol">
<div class="hr"><span class="hl">Employee Name</span>${emp.name}</div>
<div class="hr"><span class="hl">Emp ID</span>${emp.id}</div>
<div class="hr"><span class="hl">Designation</span>${emp.position||""}</div>
<div class="hr"><span class="hl">Department</span>${emp.department||""}</div>
</div></div>
<div class="wrap"><div class="ec"><table><thead><tr><th>Description</th><th class="r">Amount (SGD)</th></tr></thead><tbody>
<tr><td>Current Month Salary</td><td class="r">${fmtSGD2(p.basic)}</td></tr>
<tr><td>Housing Allowance</td><td class="r">${fmtSGD2(p.housingAllowance||0)}</td></tr>
<tr><td>Phone Allowance</td><td class="r">${fmtSGD2(p.phoneAllowance||0)}</td></tr>
<tr><td>Airfare Allowance</td><td class="r">${fmtSGD2(p.airfareAllowance||0)}</td></tr>
<tr><td>Overtime Pay</td><td class="r">${fmtSGD2(p.otPay||0)}</td></tr>
<tr><td>Leave Encashment</td><td class="r">${fmtSGD2(p.leaveEncashment||0)}</td></tr>
<tr><td>Performance Bonus</td><td class="r">${fmtSGD2(p.commission||0)}</td></tr>
<tr class="b"><td>Gross Pay</td><td class="r">${fmtSGD2(p.gross)}</td></tr>
</tbody></table></div><div class="ec"><table><thead><tr><th>Description</th><th class="r">Amount (SGD)</th></tr></thead><tbody>
<tr><td>No Paid Leave</td><td class="r">${fmtSGD2(p.nplDeduction||0)}</td></tr>
<tr><td>Income Tax</td><td class="r">${fmtSGD2(p.incomeTax||0)}</td></tr>
<tr><td>HQ Wage Deduction</td><td class="r">${fmtSGD2(p.otherDeductions||0)}</td></tr>
<tr><td>Employee CPF</td><td class="r">${fmtSGD2(p.cpfEmployee||0)}</td></tr>
<tr><td>Donation (CDAC/SINDA)</td><td class="r">${fmtSGD2(p.cdac||0)}</td></tr>
<tr class="b"><td>Total Deductions</td><td class="r">${fmtSGD2(totalDeductions)}</td></tr>
</tbody></table></div></div>
<div class="sum">
<div class="sr net"><span>Net Pay</span><span>$ ${fmtSGD2(p.netPay)}</span></div>
<div class="sr"><span>Gross Pay</span><span>$ ${fmtSGD2(p.gross)}</span></div>
<div class="sr"><span>CPF Wage</span><span>$ ${fmtSGD2(p.basic)}</span></div>
<div class="sr"><span>Employer CPF</span><span>$ ${fmtSGD2(p.cpfEmployer||0)}</span></div>
<div class="sr"><span>Employee CPF</span><span>$ ${fmtSGD2(p.cpfEmployee||0)}</span></div>
</div>
<div class="ft">This is a computer generated payslip, no signature is required.</div>
</body></html>`;
  const w=window.open("","_blank");
  if(w){w.document.write(html);w.document.close();}
  else alert("Please allow pop-ups to download payslip.");
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: PROJECTS / DEPARTMENT MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════
// ─── BENEFIT TRACKING DATA ────────────────────────────────────────────────────
const INIT_BENEFIT_TYPES = [
  {id:"bt1",name:"TCM / Traditional Medicine",icon:"🌿",limit:500,period:"annual",description:"Acupuncture, TCM, massage therapy",active:true},
  {id:"bt2",name:"Dental & Optical",icon:"🦷",limit:300,period:"annual",description:"Dental checkups, spectacles, contact lenses",active:true},
  {id:"bt3",name:"Wellness & Gym",icon:"💪",limit:600,period:"annual",description:"Gym membership, fitness classes, sports equipment",active:true},
  {id:"bt4",name:"Computer Accessories",icon:"💻",limit:500,period:"annual",description:"Mouse, keyboard, webcam, headset",active:true},
  {id:"bt5",name:"Professional Development",icon:"📚",limit:1000,period:"annual",description:"Courses, certifications, books, conferences",active:true},
  {id:"bt6",name:"Medical Claims",icon:"🏥",limit:500,period:"annual",description:"GP visit, specialist, medication",active:true},
];

// ─── CONFIGURABLE ALLOWANCE TYPES ────────────────────────────────────────────
const INIT_ALLOWANCE_TYPES = [
  {id:"al1",key:"phoneAllowance",label:"Phone Allowance",icon:"📱",active:true,cpf:false},
  {id:"al2",key:"housingAllowance",label:"Housing Allowance",icon:"🏠",active:true,cpf:false},
  {id:"al3",key:"airfareAllowance",label:"Airfare Allowance",icon:"✈️",active:true,cpf:false},
  {id:"al4",key:"transportAllowance",label:"Transport Allowance",icon:"🚌",active:false,cpf:false},
  {id:"al5",key:"mealAllowance",label:"Meal Allowance",icon:"🍱",active:false,cpf:false},
  {id:"al6",key:"otherAllowance",label:"Other Allowance",icon:"💰",active:true,cpf:false},
];

// ─── CONFIGURABLE EMPLOYEE FIELDS ────────────────────────────────────────────
const INIT_EMP_FIELD_CONFIG = [
  {section:"Personal",fields:[
    {key:"fin",label:"FIN / NRIC",type:"text",active:true,required:true},
    {key:"dob",label:"Date of Birth",type:"date",active:true,required:false},
    {key:"gender",label:"Gender",type:"text",active:true,required:false},
    {key:"nationality",label:"Nationality",type:"text",active:true,required:false},
    {key:"marital",label:"Marital Status",type:"text",active:false,required:false},
    {key:"religion",label:"Religion",type:"text",active:false,required:false},
    {key:"qualification",label:"Qualification",type:"text",active:true,required:false},
  ]},
  {section:"Contact",fields:[
    {key:"mobile",label:"Mobile No.",type:"text",active:true,required:false},
    {key:"personalEmail",label:"Personal Email",type:"email",active:true,required:false},
    {key:"workEmail",label:"Work Email",type:"email",active:true,required:false},
    {key:"residenceAddress",label:"Residence Address",type:"text",active:true,required:false},
  ]},
  {section:"Work Pass",fields:[
    {key:"workPass",label:"Work Pass Type",type:"text",active:true,required:false},
    {key:"epSpNo",label:"Work Pass No.",type:"text",active:true,required:false},
    {key:"workPassIssueDate",label:"WP Issue Date",type:"date",active:true,required:false},
    {key:"workPassExpiryDate",label:"WP Expiry Date",type:"date",active:true,required:false},
    {key:"passportNumber",label:"Passport No.",type:"text",active:true,required:false},
    {key:"passportIssueDate",label:"Passport Issue",type:"date",active:true,required:false},
    {key:"passportExpiryDate",label:"Passport Expiry",type:"date",active:true,required:false},
    {key:"probationEndDate",label:"Probation End Date",type:"date",active:true,required:false},
  ]},
  {section:"Bank & Payroll",fields:[
    {key:"bankName",label:"Bank Name",type:"text",active:true,required:false},
    {key:"bankAccount",label:"Bank Account No.",type:"text",active:true,required:false},
  ]},
  {section:"Emergency Contact",fields:[
    {key:"emergencyContactName",label:"Emergency Contact Name",type:"text",active:true,required:false},
    {key:"emergencyContactPhone",label:"Emergency Contact Phone",type:"text",active:true,required:false},
    {key:"emergencyContactRelation",label:"Relationship",type:"text",active:true,required:false},
  ]},
];

// ─── EMPLOYEE: BENEFITS (My Benefit Claims) ──────────────────────────────────
function EmpBenefits({empId,benefitTypes,benefitClaims,onAdd,lang}){
  const safeBT=Array.isArray(benefitTypes)&&benefitTypes.length?benefitTypes:INIT_BENEFIT_TYPES;
  const safeClaims=Array.isArray(benefitClaims)?benefitClaims:[];
  const myClaims=safeClaims.filter(c=>c.empId===empId);
  const year=new Date().getFullYear();
  const [showForm,setShowForm]=useState(false);
  const [selType,setSelType]=useState("");
  const [form,setForm]=useState({benefitTypeId:"",date:todayStr(),amount:"",description:"",receipt:""});
  const setFk=(k,v)=>setForm(p=>({...p,[k]:v}));
  const fileRef=useRef();
  function handleFile(e){const f=e.target.files?.[0];if(!f)return;const r=new FileReader();r.onload=ev=>setFk("receipt",ev.target.result);r.readAsDataURL(f);}
  function getUsed(btId){return myClaims.filter(c=>c.benefitTypeId===btId&&c.status==="Approved"&&c.date?.startsWith(String(year))).reduce((s,c)=>s+Number(c.amount||0),0);}
  function getPending(btId){return myClaims.filter(c=>c.benefitTypeId===btId&&c.status==="Pending"&&c.date?.startsWith(String(year))).reduce((s,c)=>s+Number(c.amount||0),0);}
  function submit(){if(!form.benefitTypeId||!form.amount||!form.description.trim())return;onAdd({...form,empId,amount:Number(form.amount),status:"Pending",submittedDate:todayStr(),adminComment:""});setShowForm(false);setForm({benefitTypeId:"",date:todayStr(),amount:"",description:"",receipt:""});}
  return<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <PageTitle title="💊 My Benefits" sub={`${year} benefit entitlements and claims`} actions={[<Btn onClick={()=>setShowForm(true)}>+ Submit Claim</Btn>]}/>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
      {safeBT.filter(bt=>bt.active).map(bt=>{
        const used=getUsed(bt.id);const pending=getPending(bt.id);const limit=Number(bt.limit)||0;const pct=limit>0?Math.min(100,(used/limit)*100):0;
        return<Card key={bt.id} style={{padding:"16px 18px"}}>
          <div style={{display:"flex",gap:12,alignItems:"flex-start",marginBottom:12}}>
            <div style={{width:44,height:44,borderRadius:12,background:C.accentL,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{bt.icon||"💰"}</div>
            <div style={{flex:1}}><div style={{fontWeight:700,fontSize:14}}>{bt.name}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>{bt.description}</div></div>
          </div>
          <div style={{marginBottom:8}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:5}}>
              <span style={{color:C.muted}}>Used: <strong style={{color:C.danger}}>{sgd(used)}</strong></span>
              <span style={{color:C.muted}}>Limit: <strong>{sgd(limit)}</strong></span>
            </div>
            <div style={{height:7,background:C.border,borderRadius:99,overflow:"hidden"}}>
              <div style={{width:`${pct}%`,height:"100%",background:pct>=90?C.danger:pct>=70?C.warning:C.success,borderRadius:99,transition:"width 0.5s ease"}}/>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginTop:4}}>
              <span style={{color:pending>0?C.warning:C.muted}}>{pending>0?`⏳ ${sgd(pending)} pending`:""}</span>
              <span style={{fontWeight:700,color:pct>=90?C.danger:C.success}}>Balance: {sgd(Math.max(0,limit-used))}</span>
            </div>
          </div>
          <Btn sm v="outline" onClick={()=>{setShowForm(true);setFk("benefitTypeId",bt.id);}}>+ Claim This</Btn>
        </Card>;
      })}
    </div>
    {myClaims.length>0&&<><SecTitle icon="📋">My Claim History</SecTitle>
    <TTable cols={["Benefit","Date","Amount","Description","Status","Submitted"]}
      rows={[...myClaims].sort((a,b)=>b.submittedDate?.localeCompare(a.submittedDate||"")||0).map(c=>{const bt=safeBT.find(x=>x.id===c.benefitTypeId);return<TR key={c.id}>
        <TD><span style={{fontSize:18,marginRight:6}}>{bt?.icon||"💰"}</span><span style={{fontWeight:600,fontSize:12}}>{bt?.name||c.benefitTypeId}</span></TD>
        <TD style={{fontSize:12}}>{fmtDate(c.date)}</TD>
        <TD style={{fontWeight:700,color:C.success}}>{sgd(c.amount)}</TD>
        <TD style={{fontSize:12,color:C.muted}}>{c.description}</TD>
        <TD><Badge s={c.status}/></TD>
        <TD style={{fontSize:11,color:C.muted}}>{fmtDate(c.submittedDate)}</TD>
      </TR>;})}/>
    </>}
    {showForm&&<Modal title="💊 Submit Benefit Claim" onClose={()=>setShowForm(false)} width={480}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Sel label="Benefit Type" value={form.benefitTypeId} onChange={v=>setFk("benefitTypeId",v)} required options={safeBT.filter(bt=>bt.active).map(bt=>({v:bt.id,l:`${bt.icon||""} ${bt.name} (limit: ${sgd(bt.limit)})`}))}/>
        <Grid><Inp label="Date" type="date" value={form.date} onChange={v=>setFk("date",v)}/><Inp label="Amount (S$)" type="number" value={form.amount} onChange={v=>setFk("amount",v)} required/></Grid>
        <Inp label="Description / Purpose" value={form.description} onChange={v=>setFk("description",v)} rows={2} placeholder="e.g. Dental checkup at Bright Dental Clinic"/>
        <div><label style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:6}}>Receipt / Invoice</label>
          <Btn v="ghost" sm onClick={()=>fileRef.current?.click()}>📎 Attach Receipt</Btn>
          <input ref={fileRef} type="file" accept="image/*,application/pdf" style={{display:"none"}} onChange={handleFile}/>
          {form.receipt&&<div style={{fontSize:11,color:C.success,marginTop:4}}>✓ File attached</div>}
        </div>
        <div style={{background:C.accentL,borderRadius:10,padding:"10px 14px",fontSize:12,color:C.accentT}}>📌 Claims require admin approval. Allow 3-5 working days for processing.</div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowForm(false)}>Cancel</Btn><Btn onClick={submit}>Submit Claim</Btn></div>
      </div>
    </Modal>}
  </div>;
}

// ─── ADMIN: BENEFITS SETUP & APPROVAL ────────────────────────────────────────
function AdminBenefits({benefitTypes,benefitClaims,employees,onAddType,onEditType,onDeleteType,onApprove,onReject,lang}){
  const safeBT=Array.isArray(benefitTypes)&&benefitTypes.length?benefitTypes:INIT_BENEFIT_TYPES;
  const safeClaims=Array.isArray(benefitClaims)?benefitClaims:[];
  const [tab,setTab]=useState("claims");
  const [showTypeForm,setShowTypeForm]=useState(false);
  const [editType,setEditType]=useState(null);
  const [tf,setTf]=useState({name:"",icon:"💰",limit:"",period:"annual",description:"",active:true});
  const [filter,setFilter]=useState("Pending");
  const setTfk=(k,v)=>setTf(p=>({...p,[k]:v}));
  function saveType(){if(!tf.name.trim())return;const d={...tf,limit:Number(tf.limit)||0};editType?onEditType(editType.id,d):onAddType({...d,id:"bt"+Date.now()});setShowTypeForm(false);}
  function totalClaimed(btId){return safeClaims.filter(c=>c.benefitTypeId===btId&&c.status==="Approved").reduce((s,c)=>s+Number(c.amount||0),0);}
  const filtered=filter==="All"?safeClaims:safeClaims.filter(c=>c.status===filter);
  const totalPending=safeClaims.filter(c=>c.status==="Pending").reduce((s,c)=>s+Number(c.amount||0),0);
  const TABS=[{id:"claims",label:"Claims",icon:"📋"},{id:"types",label:"Benefit Types",icon:"⚙️"}];
  return<div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title="💊 Benefits Management" sub={`${safeClaims.filter(c=>c.status==="Pending").length} pending approvals · ${sgd(totalPending)} pending`}/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="claims"&&<>
      <FilterBar>{["All","Pending","Approved","Rejected"].map(s=><StatusPill key={s} s={s} active={filter===s} onClick={setFilter}/>)}</FilterBar>
      <TTable cols={["Employee","Benefit","Date","Amount","Description","Status","Actions"]}
        rows={filtered.sort((a,b)=>b.submittedDate?.localeCompare(a.submittedDate||"")||0).map(c=>{const emp=employees.find(e=>e.id===c.empId);const bt=safeBT.find(x=>x.id===c.benefitTypeId);return<TR key={c.id}>
          <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={emp?.name||"?"} size={26}/><span style={{fontWeight:600,fontSize:12}}>{emp?.name||c.empId}</span></div></TD>
          <TD style={{fontSize:12}}>{bt?.icon||"💰"} {bt?.name||"—"}</TD>
          <TD style={{fontSize:12}}>{fmtDate(c.date)}</TD>
          <TD style={{fontWeight:700,color:C.success}}>{sgd(c.amount)}</TD>
          <TD style={{fontSize:12,color:C.muted,maxWidth:150}}>{c.description?.slice(0,40)||"—"}</TD>
          <TD><Badge s={c.status}/></TD>
          <TD><div style={{display:"flex",gap:4}}>
            {c.status==="Pending"&&<><Btn v="success" sm onClick={()=>onApprove(c.id)}>✓</Btn><Btn v="danger" sm onClick={()=>onReject(c.id)}>✕</Btn></>}
            {c.receipt&&<Btn v="ghost" sm onClick={()=>{const w=window.open("","_blank");w&&(w.document.write(`<img src="${c.receipt}" style="max-width:100%"/>`),w.document.close());}}>📎</Btn>}
          </div></TD>
        </TR>;})}/>
    </>}
    {tab==="types"&&<>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:8}}><Btn onClick={()=>{setEditType(null);setTf({name:"",icon:"💰",limit:"",period:"annual",description:"",active:true});setShowTypeForm(true);}}>+ Add Benefit Type</Btn></div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:10}}>
        {safeBT.map(bt=><Card key={bt.id} style={{opacity:bt.active?1:0.6}}>
          <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:8}}>
            <span style={{fontSize:24}}>{bt.icon||"💰"}</span>
            <div style={{flex:1}}><div style={{fontWeight:700}}>{bt.name}</div><div style={{fontSize:11,color:C.muted}}>{bt.period} · Limit: {sgd(bt.limit)}</div></div>
            <Badge s={bt.active?"Active":"Inactive"}/>
          </div>
          <div style={{fontSize:12,color:C.muted,marginBottom:10}}>{bt.description}</div>
          <div style={{fontSize:12,fontWeight:700,color:C.accent,marginBottom:10}}>Total claimed: {sgd(totalClaimed(bt.id))}</div>
          <div style={{display:"flex",gap:6}}>
            <Btn v="outline" sm onClick={()=>{setEditType(bt);setTf({name:bt.name,icon:bt.icon||"💰",limit:String(bt.limit),period:bt.period,description:bt.description,active:bt.active});setShowTypeForm(true);}}>Edit</Btn>
            <Btn v={bt.active?"warning":"success"} sm onClick={()=>onEditType(bt.id,{...bt,active:!bt.active})}>{bt.active?"Disable":"Enable"}</Btn>
          </div>
        </Card>)}
      </div>
      {showTypeForm&&<Modal title={editType?"Edit Benefit Type":"Add Benefit Type"} onClose={()=>setShowTypeForm(false)} width={460}>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <Grid cols={3}><Inp label="Icon (emoji)" value={tf.icon} onChange={v=>setTfk("icon",v)} placeholder="💰"/><div style={{gridColumn:"span 2"}}><Inp label="Benefit Name" value={tf.name} onChange={v=>setTfk("name",v)} required/></div></Grid>
          <Grid><Inp label="Annual Limit (S$)" type="number" value={tf.limit} onChange={v=>setTfk("limit",v)}/><Sel label="Period" value={tf.period} onChange={v=>setTfk("period",v)} options={["annual","monthly"]}/></Grid>
          <Inp label="Description" value={tf.description} onChange={v=>setTfk("description",v)} rows={2}/>
          <label style={{display:"flex",gap:8,fontSize:13,cursor:"pointer"}}><input type="checkbox" checked={tf.active} onChange={e=>setTfk("active",e.target.checked)}/>Active (visible to employees)</label>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowTypeForm(false)}>Cancel</Btn><Btn onClick={saveType}>Save</Btn></div>
        </div>
      </Modal>}
    </>}
  </div>;
}

// ─── ADMIN: FIELD CONFIG (Settings tab) ──────────────────────────────────────
function FieldConfigManager({fieldConfig,onUpdate}){
  const safe=Array.isArray(fieldConfig)&&fieldConfig.length?fieldConfig:INIT_EMP_FIELD_CONFIG;
  function toggle(secIdx,fIdx){
    const updated=safe.map((sec,si)=>si!==secIdx?sec:{...sec,fields:sec.fields.map((f,fi)=>fi!==fIdx?f:{...f,active:!f.active})});
    onUpdate(updated);
  }
  return<div>
    <div style={{fontSize:13,color:C.muted,marginBottom:14}}>Toggle which fields appear in the Employee form. Core fields (ID, Name, Department) are always shown.</div>
    {safe.map((sec,si)=><div key={sec.section} style={{marginBottom:16}}>
      <SecTitle icon="📋">{sec.section}</SecTitle>
      <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
        {(sec.fields||[]).map((f,fi)=><button key={f.key} onClick={()=>toggle(si,fi)} style={{display:"flex",gap:6,alignItems:"center",background:f.active?C.successL:C.bg,border:`1.5px solid ${f.active?C.success:C.border}`,borderRadius:9,padding:"5px 12px",cursor:"pointer",fontSize:12,fontWeight:600,color:f.active?C.successD:C.muted,fontFamily:"inherit"}}>
          <span>{f.active?"✓":"○"}</span>{f.label}
        </button>)}
      </div>
    </div>)}
  </div>;
}

// ─── ADMIN: ALLOWANCE TYPE MANAGER ───────────────────────────────────────────
function AllowanceTypeManager({allowanceTypes,onUpdate}){
  const safe=Array.isArray(allowanceTypes)&&allowanceTypes.length?allowanceTypes:INIT_ALLOWANCE_TYPES;
  const [showAdd,setShowAdd]=useState(false);
  const [nf,setNf]=useState({key:"",label:"",icon:"💰",active:true,cpf:false});
  function toggle(id){onUpdate(safe.map(a=>a.id===id?{...a,active:!a.active}:a));}
  function addNew(){if(!nf.label.trim())return;const key=nf.label.toLowerCase().replace(/\s+/g,"_")+"Allowance";onUpdate([...safe,{...nf,key,id:"al"+Date.now()}]);setShowAdd(false);setNf({key:"",label:"",icon:"💰",active:true,cpf:false});}
  function del(id){onUpdate(safe.filter(a=>a.id!==id));}
  return<div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
      <div style={{fontSize:13,color:C.muted}}>Configure which allowance types appear in payroll processing.</div>
      <Btn sm onClick={()=>setShowAdd(true)}>+ Add Allowance Type</Btn>
    </div>
    <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:12}}>
      {safe.map(a=><div key={a.id} style={{display:"flex",gap:6,alignItems:"center",background:a.active?C.accentL:C.bg,border:`1.5px solid ${a.active?C.accent:C.border}`,borderRadius:99,padding:"5px 6px 5px 12px"}}>
        <span style={{fontSize:13}}>{a.icon||"💰"}</span><span style={{fontSize:12,fontWeight:600,color:a.active?C.accentT:C.muted}}>{a.label}</span>
        <button onClick={()=>toggle(a.id)} style={{background:a.active?C.accent:C.border,color:"#fff",border:"none",borderRadius:"50%",width:18,height:18,fontSize:10,cursor:"pointer",fontFamily:"inherit"}}>{a.active?"✓":"○"}</button>
        {!a.required&&<button onClick={()=>del(a.id)} style={{background:C.dangerL,color:C.danger,border:"none",borderRadius:"50%",width:18,height:18,fontSize:10,cursor:"pointer"}}>✕</button>}
      </div>)}
    </div>
    {showAdd&&<div style={{background:C.bg,borderRadius:12,padding:"14px",border:`1px solid ${C.border}`}}>
      <Grid cols={3}><Inp label="Icon" value={nf.icon} onChange={v=>setNf(p=>({...p,icon:v}))}/><div style={{gridColumn:"span 2"}}><Inp label="Allowance Name" value={nf.label} onChange={v=>setNf(p=>({...p,label:v}))} placeholder="e.g. Cloud / WFH Allowance"/></div></Grid>
      <div style={{display:"flex",gap:10,marginTop:10,justifyContent:"flex-end"}}><Btn v="ghost" sm onClick={()=>setShowAdd(false)}>Cancel</Btn><Btn sm onClick={addNew}>Add</Btn></div>
    </div>}
  </div>;
}

const INIT_PROJECTS_DATA=[
  {id:"P001",name:"Tuas Industrial Project",code:"TIS-2026",client:"JTC Corporation",startDate:"2026-01-01",endDate:"2027-12-31",status:"Active",supervisorId:"C006",hodId:"C001",members:["C004","C005"],description:"Industrial building construction at Tuas West"},
  {id:"P002",name:"Jurong Residential Development",code:"JRD-2026",client:"HDB",startDate:"2026-03-01",endDate:"2028-06-30",status:"Active",supervisorId:"C002",hodId:"C001",members:["C005","C012"],description:"Public housing development at Jurong East"},
];

function AdminProjects({projects,employees,onAdd,onEdit,onDelete}){
  const safeProjects=Array.isArray(projects)?projects:[];
  const safeEmps=Array.isArray(employees)?employees:[];
  const [showForm,setShowForm]=useState(false);
  const [editP,setEditP]=useState(null);
  const [viewP,setViewP]=useState(null);
  const [confirmId,setConfirmId]=useState(null);
  const EF={name:"",code:"",client:"",startDate:"",endDate:"",status:"Active",supervisorId:"",hodId:"",members:[],description:""};
  const [f,setF]=useState(EF);
  const setFk=(k,v)=>setF(p=>({...p,[k]:v}));
  function openNew(){setEditP(null);setF(EF);setShowForm(true);}
  function openEdit(p){setEditP(p);setF({name:p.name||"",code:p.code||"",client:p.client||"",startDate:p.startDate||"",endDate:p.endDate||"",status:p.status||"Active",supervisorId:p.supervisorId||"",hodId:p.hodId||"",members:Array.isArray(p.members)?p.members:[],description:p.description||""});setShowForm(true);}
  function save(){if(!f.name.trim())return;editP?onEdit(editP.id,{...f,id:editP.id}):onAdd({...f,id:"P"+String(safeProjects.length+1).padStart(3,"0")});setShowForm(false);}
  function toggleMember(id){setF(p=>({...p,members:(p.members||[]).includes(id)?(p.members||[]).filter(x=>x!==id):[...(p.members||[]),id]}))}
  const statusC={Active:C.success,Completed:C.accent,OnHold:C.warning,Cancelled:C.danger};
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title="Projects" sub={`${safeProjects.length} projects`} actions={[<Btn onClick={openNew}>+ New Project</Btn>]}/>
    {safeProjects.length===0&&<Card><p style={{color:C.muted,textAlign:"center",padding:32,fontSize:13}}>No projects yet. Click + New Project to create one.</p></Card>}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      {safeProjects.map(p=>{
        const sup=safeEmps.find(e=>e.id===p.supervisorId);
        const hod=safeEmps.find(e=>e.id===p.hodId);
        const membs=safeEmps.filter(e=>(p.members||[]).includes(e.id));
        return <Card key={p.id} style={{borderLeft:`4px solid ${statusC[p.status]||C.accent}`}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
            <div><div style={{fontWeight:800,fontSize:14}}>{p.name}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>{p.code||""}{p.client?` · ${p.client}`:""}</div></div>
            <Badge s={p.status||"Active"}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,fontSize:12,marginBottom:10}}>
            <div><div style={{fontSize:10,color:C.muted}}>Supervisor</div><div style={{fontWeight:600}}>{sup?.name?.split(",")[0]||"—"}</div></div>
            <div><div style={{fontSize:10,color:C.muted}}>HOD</div><div style={{fontWeight:600}}>{hod?.name?.split(",")[0]||"—"}</div></div>
            <div><div style={{fontSize:10,color:C.muted}}>Start</div><div>{fmtDate(p.startDate)}</div></div>
            <div><div style={{fontSize:10,color:C.muted}}>End</div><div>{fmtDate(p.endDate)}</div></div>
          </div>
          <div style={{display:"flex",gap:4,flexWrap:"wrap",marginBottom:10}}>
            {membs.slice(0,5).map(m=><div key={m.id} title={m.name}><Avatar name={m.name} size={26}/></div>)}
            {membs.length>5&&<div style={{width:26,height:26,borderRadius:"50%",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:C.muted}}>+{membs.length-5}</div>}
            {membs.length===0&&<span style={{fontSize:11,color:C.muted}}>No members yet</span>}
          </div>
          <div style={{display:"flex",gap:6}}>
            <Btn v="outline" sm onClick={()=>setViewP(p)}>View</Btn>
            <Btn v="ghost" sm onClick={()=>openEdit(p)}>Edit</Btn>
            <Btn v="danger" sm onClick={()=>setConfirmId(p.id)}>Del</Btn>
          </div>
        </Card>;
      })}
    </div>
    {viewP&&<Modal title={viewP.name} onClose={()=>setViewP(null)} width={560}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
        {[["Code",viewP.code||"—"],["Client",viewP.client||"—"],["Status",viewP.status],["Start",fmtDate(viewP.startDate)],["End",fmtDate(viewP.endDate)],["HOD",safeEmps.find(e=>e.id===viewP.hodId)?.name||"—"],["Supervisor",safeEmps.find(e=>e.id===viewP.supervisorId)?.name||"—"]].map(([k,v])=><div key={k} style={{background:C.bg,borderRadius:7,padding:"8px 12px"}}><div style={{fontSize:10,color:C.muted}}>{k}</div><div style={{fontWeight:600,fontSize:13}}>{v}</div></div>)}
      </div>
      {viewP.description&&<div style={{background:C.bg,borderRadius:7,padding:"10px 12px",marginBottom:12,fontSize:13}}>{viewP.description}</div>}
      <SecTitle>Team Members ({safeEmps.filter(e=>(viewP.members||[]).includes(e.id)).length})</SecTitle>
      {safeEmps.filter(e=>(viewP.members||[]).includes(e.id)).map(e=><div key={e.id} style={{display:"flex",gap:10,alignItems:"center",background:C.bg,borderRadius:7,padding:"8px 12px",marginBottom:6}}><Avatar name={e.name} size={30}/><div><div style={{fontWeight:600,fontSize:13}}>{e.name}</div><div style={{fontSize:11,color:C.muted}}>{e.position}</div></div></div>)}
    </Modal>}
    {showForm&&<Modal title={editP?"Edit Project":"New Project"} onClose={()=>setShowForm(false)} width={660}>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Grid><Inp label="Project Name" value={f.name} onChange={v=>setFk("name",v)} required/><Inp label="Project Code" value={f.code} onChange={v=>setFk("code",v)} placeholder="e.g. TIS-2026"/></Grid>
        <Grid><Inp label="Client" value={f.client} onChange={v=>setFk("client",v)}/><Sel label="Status" value={f.status} onChange={v=>setFk("status",v)} options={["Active","Completed","OnHold","Cancelled"]}/></Grid>
        <Grid><Inp label="Start Date" type="date" value={f.startDate} onChange={v=>setFk("startDate",v)}/><Inp label="End Date" type="date" value={f.endDate} onChange={v=>setFk("endDate",v)}/></Grid>
        <Grid><Sel label="HOD (Head of Department)" value={f.hodId} onChange={v=>setFk("hodId",v)} options={[{v:"",l:"— Select —"},...safeEmps.map(e=>({v:e.id,l:`${e.id} - ${e.name}`}))]}/><Sel label="Supervisor / Reporting Officer" value={f.supervisorId} onChange={v=>setFk("supervisorId",v)} options={[{v:"",l:"— Select —"},...safeEmps.map(e=>({v:e.id,l:`${e.id} - ${e.name}`}))]}/></Grid>
        <Inp label="Description" value={f.description} onChange={v=>setFk("description",v)} rows={2}/>
        <div><label style={{fontSize:12,fontWeight:600,color:C.muted,display:"block",marginBottom:8}}>Team Members</label>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
            {safeEmps.map(e=>{const sel=(f.members||[]).includes(e.id);return<button key={e.id} onClick={()=>toggleMember(e.id)} style={{display:"flex",gap:6,alignItems:"center",background:sel?C.accentL:C.bg,border:`1px solid ${sel?C.accent:C.border}`,borderRadius:20,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:sel?700:400,color:sel?C.accent:C.muted,fontFamily:"inherit"}}>
              <Avatar name={e.name} size={18}/>{e.name.split(/[\s,]+/)[0]}
            </button>;})}
          </div>
        </div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><Btn v="ghost" onClick={()=>setShowForm(false)}>Cancel</Btn><Btn onClick={save}>Save Project</Btn></div>
      </div>
    </Modal>}
    {confirmId&&<Confirm msg="Delete this project?" onOk={()=>{onDelete(confirmId);setConfirmId(null);}} onCancel={()=>setConfirmId(null)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: TRACKER — Work Pass + Passport + Birthdays + Leave Balance
// ═══════════════════════════════════════════════════════════════════════════════
function AdminTracker({employees,leaves,leaveTypes}){
  const safeEmps=Array.isArray(employees)?employees:[];
  const safeLeaves=Array.isArray(leaves)?leaves:[];
  const safeLT=Array.isArray(leaveTypes)&&leaveTypes.length?leaveTypes:[{name:"Annual Leave",days:14},{name:"Medical Leave",days:14}];
  const [tab,setTab]=useState("workpass");
  const today=new Date();
  const wpRows=safeEmps.filter(e=>e.workPassExpiryDate||e.epExpiry).map(e=>{
    const exp=e.workPassExpiryDate||e.epExpiry||"";
    const d=daysUntil(exp);
    return{...e,_exp:exp,daysLeft:d,urgent:d!==null&&d<=30&&d>=0,warn:d!==null&&d>30&&d<=90};
  }).sort((a,b)=>(a.daysLeft===null?9999:a.daysLeft)-(b.daysLeft===null?9999:b.daysLeft));
  const ppRows=safeEmps.filter(e=>e.passportExpiryDate||e.passportExpiry).map(e=>{
    const exp=e.passportExpiryDate||e.passportExpiry||"";
    const d=daysUntil(exp);
    return{...e,_exp:exp,daysLeft:d,urgent:d!==null&&d<=60&&d>=0,warn:d!==null&&d>60&&d<=180};
  }).sort((a,b)=>(a.daysLeft===null?9999:a.daysLeft)-(b.daysLeft===null?9999:b.daysLeft));
  const bdRows=safeEmps.filter(e=>{if(!e.dob)return false;const d=new Date(e.dob);return!isNaN(d.getTime());}).map(e=>{
    const dob=new Date(e.dob);const next=new Date(today.getFullYear(),dob.getMonth(),dob.getDate());
    if(next<today)next.setFullYear(today.getFullYear()+1);
    const d=Math.ceil((next-today)/86400000);
    return{...e,daysLeft:d,nextBday:next,thisMonth:dob.getMonth()===today.getMonth()};
  }).sort((a,b)=>a.daysLeft-b.daysLeft);
  const lbRows=safeEmps.map(e=>{
    const approved=safeLeaves.filter(l=>l.empId===e.id&&l.status==="Approved");
    const alEnt=Number(e.annualLeave)||14;
    const alUsed=approved.filter(l=>l.type==="Annual Leave").reduce((s,l)=>s+(Number(l.days)||0),0);
    const mlUsed=approved.filter(l=>l.type==="Medical Leave").reduce((s,l)=>s+(Number(l.days)||0),0);
    return{...e,alEnt,alUsed,alBal:alEnt-alUsed,mlUsed,mlBal:14-mlUsed};
  }).sort((a,b)=>a.alBal-b.alBal);
  const TABS=[{id:"workpass",label:"Work Pass"},{id:"passport",label:"Passport"},{id:"birthday",label:"Birthdays"},{id:"leaveBal",label:"Leave Balance"}];
  function doExport(){
    if(tab==="workpass")exportXLS(wpRows.map(e=>({ID:e.id,Name:e.name,"Work Pass":e.workPass,"WP No.":e.epSpNo||"","Expiry":fmtDate(e._exp),"Days Left":e.daysLeft})),"WorkPass","workpass.xlsx");
    else if(tab==="birthday")exportXLS(bdRows.map(e=>({ID:e.id,Name:e.name,DOB:fmtDate(e.dob),Age:calcAge(e.dob),"Days Away":e.daysLeft,"This Month":e.thisMonth?"Yes":"No"})),"Birthdays","birthdays.xlsx");
    else exportXLS(lbRows.map(e=>({ID:e.id,Name:e.name,Dept:e.department,"AL Entitlement":e.alEnt,"AL Used":e.alUsed,"AL Balance":e.alBal,"ML Used":e.mlUsed,"ML Balance":e.mlBal})),"LeaveBalance","leave_balance.xlsx");
  }
  return <div style={{display:"flex",flexDirection:"column",gap:14}}>
    <PageTitle title="Tracker" actions={[<Btn v="ghost" onClick={doExport}>Export Excel</Btn>]}/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="workpass"&&<>
      <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
        <StatCard label="Expiring ≤30 days" value={wpRows.filter(r=>r.urgent).length} icon="!" color={C.danger}/>
        <StatCard label="Expiring 31–90 days" value={wpRows.filter(r=>r.warn).length} icon="!" color={C.warning}/>
        <StatCard label="OK" value={wpRows.filter(r=>!r.urgent&&!r.warn).length} icon="✓" color={C.success}/>
      </div>
      {wpRows.length===0&&<Card><p style={{color:C.muted,textAlign:"center",padding:24}}>No work pass expiry dates entered. Update employees' Work Pass Expiry Date field.</p></Card>}
      <TTable cols={["Employee","Work Pass","WP No.","Expiry Date","Days Left","Alert"]}
        rows={wpRows.map(e=><TR key={e.id}>
          <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={e.name} size={26}/><div><div style={{fontWeight:600,fontSize:12}}>{e.name}</div><div style={{fontSize:10,color:C.muted}}>{e.id}</div></div></div></TD>
          <TD><Badge s={e.workPass||"SC"}/></TD><TD style={{fontSize:12}}>{e.epSpNo||"—"}</TD>
          <TD style={{fontSize:12}}>{fmtDate(e._exp)}</TD>
          <TD><span style={{fontWeight:800,color:e.urgent?C.danger:e.warn?C.warning:C.success}}>{e.daysLeft===0?"TODAY":e.daysLeft!==null?e.daysLeft+"d":"—"}</span></TD>
          <TD>{e.urgent?<Badge s="Rejected"/>:e.warn?<Badge s="Pending"/>:<Badge s="Active"/>}</TD>
        </TR>)}/>
    </>}
    {tab==="passport"&&<>
      {ppRows.length===0&&<Card><p style={{color:C.muted,textAlign:"center",padding:24}}>No passport expiry dates entered. Update employees' Passport Expiry Date field.</p></Card>}
      <TTable cols={["Employee","Nationality","Passport No.","Expiry Date","Days Left","Alert"]}
        rows={ppRows.map(e=><TR key={e.id}>
          <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={e.name} size={26}/><div><div style={{fontWeight:600,fontSize:12}}>{e.name}</div><div style={{fontSize:10,color:C.muted}}>{e.nationality}</div></div></div></TD>
          <TD style={{fontSize:12}}>{e.nationality}</TD>
          <TD style={{fontSize:12}}>{e.passportNumber||e.passportNo||"—"}</TD>
          <TD style={{fontSize:12}}>{fmtDate(e._exp)}</TD>
          <TD><span style={{fontWeight:800,color:e.urgent?C.danger:e.warn?C.warning:C.success}}>{e.daysLeft===0?"TODAY":e.daysLeft!==null?e.daysLeft+"d":"—"}</span></TD>
          <TD>{e.urgent?<Badge s="Rejected"/>:e.warn?<Badge s="Pending"/>:<Badge s="Active"/>}</TD>
        </TR>)}/>
    </>}
    {tab==="birthday"&&<>
      <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
        <StatCard label="This Month" value={bdRows.filter(r=>r.thisMonth).length} icon="🎂" color={C.purple}/>
        <StatCard label="Next 7 Days" value={bdRows.filter(r=>r.daysLeft<=7).length} icon="!" color={C.orange}/>
      </div>
      <TTable cols={["Employee","Dept","Date of Birth","Age","Next Birthday","Days Away","This Month"]}
        rows={bdRows.map(e=><TR key={e.id}>
          <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={e.name} size={26}/><div><div style={{fontWeight:600,fontSize:12}}>{e.name}</div><div style={{fontSize:10,color:C.muted}}>{e.id}</div></div></div></TD>
          <TD style={{fontSize:12}}>{e.department}</TD><TD style={{fontSize:12}}>{fmtDate(e.dob)}</TD>
          <TD style={{fontWeight:700}}>{calcAge(e.dob)}</TD>
          <TD style={{fontSize:12}}>{e.nextBday.toLocaleDateString("en-SG",{day:"numeric",month:"long",year:"numeric"})}</TD>
          <TD><span style={{fontWeight:800,color:e.daysLeft<=7?C.purple:C.muted}}>{e.daysLeft===0?"TODAY":e.daysLeft+"d"}</span></TD>
          <TD>{e.thisMonth?<span style={{color:C.purple,fontWeight:700}}>🎂 This Month</span>:<span style={{color:C.muted}}>—</span>}</TD>
        </TR>)}/>
    </>}
    {tab==="leaveBal"&&<TTable cols={["Employee","Dept","AL Entitlement","AL Used","AL Balance","ML Used","ML Balance"]}
      rows={lbRows.map(e=><TR key={e.id}>
        <TD><div style={{display:"flex",gap:8,alignItems:"center"}}><Avatar name={e.name} size={26}/><div><div style={{fontWeight:600,fontSize:12}}>{e.name}</div><div style={{fontSize:10,color:C.muted}}>{e.id}</div></div></div></TD>
        <TD style={{fontSize:12}}>{e.department}</TD>
        <TD style={{textAlign:"center",fontWeight:600}}>{e.alEnt}</TD>
        <TD style={{textAlign:"center",color:C.danger,fontWeight:600}}>{e.alUsed}</TD>
        <TD><div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{flex:1,height:6,background:C.border,borderRadius:99,overflow:"hidden"}}><div style={{width:`${Math.min(100,e.alEnt>0?(e.alUsed/e.alEnt)*100:0)}%`,height:"100%",background:e.alBal<=2?C.danger:e.alBal<=5?C.warning:C.success}}/></div>
          <span style={{fontWeight:700,color:e.alBal<=2?C.danger:e.alBal<=5?C.warning:C.success,minWidth:20,fontSize:13}}>{e.alBal}</span>
        </div></TD>
        <TD style={{textAlign:"center",color:C.danger,fontWeight:600}}>{e.mlUsed}</TD>
        <TD style={{textAlign:"center",fontWeight:700,color:e.mlBal<=3?C.warning:C.success}}>{e.mlBal}</TD>
      </TR>)}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ADMIN: DASHBOARD — Department summary + Alerts + Calendar
// ═══════════════════════════════════════════════════════════════════════════════
function AdminDashboard({employees,leaves,claims,attendance,announcements,calEvents,onNavigate,onLeaveAction,session,lang}){
  const safeEmps=Array.isArray(employees)?employees:[];
  const safeLeaves=Array.isArray(leaves)?leaves:[];
  const safeCal=Array.isArray(calEvents)?calEvents:[];
  const today=new Date();
  const pendingLeaves=safeLeaves.filter(l=>l.status==="Pending Supervisor"||l.status==="Pending HR");
  const pendingClaims=(Array.isArray(claims)?claims:[]).filter(c=>c.status==="Pending Supervisor"||c.status==="Pending HR");
  // Dept breakdown
  const depts=[...new Set(safeEmps.map(e=>e.department).filter(Boolean))].sort();
  const deptStats=depts.map(d=>({dept:d,total:safeEmps.filter(e=>e.department===d).length,active:safeEmps.filter(e=>e.department===d&&e.status==="Active").length}));
  // Upcoming holidays (next 60 days)
  const upHolidays=safeCal.filter(e=>e.type==="Holiday"&&e.date>=todayStr()).sort((a,b)=>a.date.localeCompare(b.date)).slice(0,5);
  // Work pass expiry ≤90 days
  const wpAlerts=safeEmps.filter(e=>{const d=daysUntil(e.workPassExpiryDate||e.epExpiry);return d!==null&&d<=90&&d>=0;}).sort((a,b)=>(daysUntil(a.workPassExpiryDate||a.epExpiry)||999)-(daysUntil(b.workPassExpiryDate||b.epExpiry)||999));
  // Birthdays this month
  const bdThisMonth=safeEmps.filter(e=>e.dob&&!isNaN(new Date(e.dob))&&new Date(e.dob).getMonth()===today.getMonth());
  return <div style={{display:"flex",flexDirection:"column",gap:18}}>
    <div><h1 style={{fontSize:20,fontWeight:800,color:C.text,margin:0}}>Dashboard</h1><p style={{color:C.muted,margin:"3px 0 0",fontSize:13}}>{today.toLocaleDateString("en-SG",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</p></div>
    <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
      <StatCard label="Total Employees" value={safeEmps.length} sub={`${safeEmps.filter(e=>e.status==="Active").length} active`} icon="E" color={C.accent}/>
      <StatCard label="Pending Approvals" value={pendingLeaves.length+pendingClaims.length} sub="leaves & claims" icon="P" color={C.warning}/>
      <StatCard label="WP Expiring Soon" value={wpAlerts.length} sub="within 90 days" icon="!" color={C.danger}/>
      <StatCard label="Birthdays This Month" value={bdThisMonth.length} icon="🎂" color={C.purple}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      {/* Department breakdown */}
      <Card>
        <SecTitle>Department Headcount</SecTitle>
        {deptStats.length===0&&<p style={{color:C.muted,fontSize:13}}>No department data.</p>}
        {deptStats.map(d=><div key={d.dept} style={{marginBottom:10}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
            <span style={{fontSize:12,fontWeight:600}}>{d.dept}</span>
            <span style={{fontSize:12,fontWeight:700,color:C.accent}}>{d.active}/{d.total}</span>
          </div>
          <div style={{height:6,background:C.border,borderRadius:99,overflow:"hidden"}}>
            <div style={{width:`${safeEmps.length>0?(d.total/safeEmps.length)*100:0}%`,height:"100%",background:C.accent,borderRadius:99}}/>
          </div>
        </div>)}
      </Card>
      {/* Pending approvals */}
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <SecTitle>Pending Approvals</SecTitle>
          <Btn v="outline" sm onClick={()=>onNavigate("leaves")}>View All</Btn>
        </div>
        {pendingLeaves.length===0&&pendingClaims.length===0&&<p style={{color:C.muted,fontSize:13}}>All clear — no pending items.</p>}
        {pendingLeaves.slice(0,3).map(l=>{const emp=safeEmps.find(e=>e.id===l.empId);return <div key={l.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${C.border}`}}>
          <Avatar name={emp?.name||"?"} size={30}/><div style={{flex:1}}><div style={{fontWeight:600,fontSize:12}}>{emp?.name||l.empId}</div><div style={{fontSize:11,color:C.muted}}>{l.type} · {fmtDate(l.from)}</div></div>
          <Btn v="success" sm onClick={()=>onLeaveAction(l.id,l.status==="Pending Supervisor"?"Pending HR":"Approved",session?.name||"","")}>✓</Btn>
          <Btn v="danger" sm onClick={()=>onLeaveAction(l.id,"Rejected",session?.name||"","")}>✕</Btn>
        </div>;})}
      </Card>
      {/* Work pass alerts */}
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <SecTitle>Work Pass Expiry Alerts</SecTitle>
          <Btn v="outline" sm onClick={()=>onNavigate("tracker")}>View Tracker</Btn>
        </div>
        {wpAlerts.length===0&&<p style={{color:C.muted,fontSize:13}}>No work pass expiring within 90 days.</p>}
        {wpAlerts.slice(0,5).map(e=>{const d=daysUntil(e.workPassExpiryDate||e.epExpiry);return <div key={e.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:d<=30?C.danger:C.warning,flexShrink:0}}/>
          <div style={{flex:1}}><span style={{fontWeight:600,fontSize:12}}>{e.name}</span><span style={{color:C.muted,fontSize:11}}> · {e.workPass}</span></div>
          <span style={{fontSize:12,fontWeight:800,color:d<=30?C.danger:C.warning}}>{d===0?"TODAY":d+"d"}</span>
        </div>;})}
      </Card>
      {/* Birthdays + Upcoming holidays */}
      <Card>
        <SecTitle>Birthdays This Month 🎂</SecTitle>
        {bdThisMonth.length===0&&<p style={{color:C.muted,fontSize:12,marginBottom:12}}>No birthdays this month.</p>}
        {bdThisMonth.map(e=><div key={e.id} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
          <Avatar name={e.name} size={28}/><div><div style={{fontWeight:600,fontSize:12}}>{e.name}</div><div style={{fontSize:11,color:C.purple}}>{new Date(e.dob).toLocaleDateString("en-SG",{day:"numeric",month:"short"})} · Turns {calcAge(e.dob)+1}</div></div>
        </div>)}
        <div style={{marginTop:12,paddingTop:10,borderTop:`1px solid ${C.border}`}}>
          <SecTitle>Upcoming Public Holidays</SecTitle>
          {upHolidays.length===0&&<p style={{color:C.muted,fontSize:12}}>No upcoming holidays in calendar. Add them via Calendar page.</p>}
          {upHolidays.map(h=><div key={h.id} style={{display:"flex",justifyContent:"space-between",marginBottom:6,fontSize:12}}><span style={{fontWeight:600}}>{h.title}</span><span style={{color:C.muted}}>{fmtDate(h.date)}</span></div>)}
        </div>
      </Card>
    </div>
    {/* Latest announcements */}
    <Card>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <SecTitle>Latest Announcements</SecTitle>
        <Btn v="outline" sm onClick={()=>onNavigate("memo")}>View All</Btn>
      </div>
      {(Array.isArray(announcements)?announcements:[]).slice(0,3).map(a=><div key={a.id} style={{borderLeft:`3px solid ${C.accent}`,paddingLeft:10,marginBottom:10}}>
        <div style={{fontWeight:700,fontSize:13}}>{a.title}</div>
        <div style={{fontSize:11,color:C.muted,marginTop:2}}>{fmtDate(a.date)} · {a.author||""}</div>
      </div>)}
    </Card>
  </div>;
}


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
const EMP_NAV=[{id:"dashboard",l:"dashboard",icon:"🏠"},{id:"attendance",l:"attendance",icon:"📍"},{id:"leaves",l:"leaves",icon:"🌴"},{id:"claims",l:"claims",icon:"💵"},{id:"benefits",l:"Benefits",icon:"💊"},{id:"payslip",l:"payslip",icon:"💰"},{id:"memo",l:"memo",icon:"📢"},{id:"policy",l:"policy",icon:"📋"},{id:"training",l:"training",icon:"🎓"},{id:"feedback",l:"feedback",icon:"💬"},{id:"appraisal",l:"appraisal",icon:"⭐"},{id:"calendar",l:"calendar",icon:"📅"},{id:"shift",l:"shift",icon:"🕐"}];
const ADMIN_NAV=[{id:"dashboard",l:"dashboard",icon:"🏠"},{id:"employees",l:"employees",icon:"👥"},{id:"projects",l:"Projects",icon:"🏗️"},{id:"users",l:"settings",icon:"🔑"},{id:"attendance",l:"attendance",icon:"📍"},{id:"leaves",l:"leaves",icon:"🌴"},{id:"claims",l:"claims",icon:"💵"},{id:"benefits",l:"Benefits",icon:"💊"},{id:"payroll",l:"payroll",icon:"💰"},{id:"tracker",l:"Tracker",icon:"📊"},{id:"memo",l:"memo",icon:"📢"},{id:"policy",l:"policy",icon:"📋"},{id:"training",l:"training",icon:"🎓"},{id:"feedback",l:"feedback",icon:"💬"},{id:"appraisal",l:"appraisal",icon:"⭐"},{id:"calendar",l:"calendar",icon:"📅"},{id:"shift",l:"shift",icon:"🕐"}];


// ─── SETTINGS HUB ────────────────────────────────────────────────────────────
function SettingsHub({users,employees,onAddUser,onEditUser,onDeleteUser,onResetPw,empFieldConfig,onUpdateFieldConfig,allowanceTypes,onUpdateAllowanceTypes,lang}){
  const [tab,setTab]=useState("users");
  const TABS=[{id:"users",label:"👤 User Accounts"},{id:"fields",label:"🗂️ Employee Fields"},{id:"allowances",label:"💰 Allowance Types"}];
  return<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <PageTitle title="⚙️ Settings"/>
    <TabBar tabs={TABS} active={tab} onChange={setTab}/>
    {tab==="users"&&<Card><UserMgmt users={users} employees={employees} onAdd={onAddUser} onEdit={onEditUser} onDelete={onDeleteUser} onResetPw={onResetPw} lang={lang}/></Card>}
    {tab==="fields"&&<Card><FieldConfigManager fieldConfig={empFieldConfig} onUpdate={onUpdateFieldConfig}/></Card>}
    {tab==="allowances"&&<Card><AllowanceTypeManager allowanceTypes={allowanceTypes} onUpdate={onUpdateAllowanceTypes}/></Card>}
  </div>;
}

export default function App(){
  const [session,setSession]=useState(null);
  const [lang,setLang]=useState("en");
  const [page,setPage]=useState("dashboard");
  const [toast,setToast]=useState(null);
  const [loading,setLoading]=useState(true);
  const [projects,setProjects]=useState(INIT_PROJECTS_DATA);
  const [company,setCompany]=useState(INIT_COMPANY);
  const [benefitTypes,setBenefitTypes]=useState(INIT_BENEFIT_TYPES);
  const [benefitClaims,setBenefitClaims]=useState([]);
  const [allowanceTypes,setAllowanceTypes]=useState(INIT_ALLOWANCE_TYPES);
  const [empFieldConfig,setEmpFieldConfig]=useState(INIT_EMP_FIELD_CONFIG);

  // ── Data states (populated by Firestore real-time listeners) ──────────────
  const [employees,setEmployees]=useState([]);
  const [users,setUsers]=useState([]);
  const [leaves,setLeaves]=useState([]);
  const [leaveTypes,setLeaveTypes]=useState(SG_LEAVE_TYPES);
  const [claims,setClaims]=useState([]);
  const [payroll,setPayroll]=useState([]);
  const [memos,setMemos]=useState([]);
  const [policies,setPolicies]=useState([]);
  const [trainings,setTrainings]=useState([]);
  const [feedbacks,setFeedbacks]=useState([]);
  const [feedbackSections,setFeedbackSections]=useState([]);
  const [appraisalForms,setAppraisalForms]=useState([]);
  const [appraisalSubs,setAppraisalSubs]=useState([]);
  const [attendance,setAttendance]=useState([]);
  const [shifts,setShifts]=useState([]);
  const [calEvents,setCalEvents]=useState([]);

  function toast_(msg,type){setToast({msg,type:type||"success"});}

  // ── Firestore helpers ─────────────────────────────────────────────────────
  const snap2arr=s=>s.docs.map(d=>({...d.data(),id:d.id}));
  const fsAdd=(col,data)=>addDoc(collection(db,col),data);
  const fsSet=(col,id,data)=>setDoc(doc(db,col,id),data);
  const fsUp=(col,id,data)=>updateDoc(doc(db,col,id),data);
  const fsDel=(col,id)=>deleteDoc(doc(db,col,id));
  async function fsAddId(col,data){const r=await addDoc(collection(db,col),data);await fsUp(col,r.id,{id:r.id});return r.id;}

  // ── Seed + Subscribe ──────────────────────────────────────────────────────
  useEffect(()=>{
    async function init(){
      try{
        const empSnap=await getDocs(collection(db,'employees'));
        if(empSnap.size===0){
          const batch=writeBatch(db);
          INIT_EMPLOYEES.forEach(e=>batch.set(doc(db,'employees',e.id),e));
          INIT_USERS.forEach(u=>batch.set(doc(db,'users',u.id),u));
          SG_LEAVE_TYPES.forEach((lt,i)=>batch.set(doc(db,'leaveTypes',`lt${i}`),lt));
          const seedCol=(items,col)=>items.forEach(item=>{
            const r=doc(collection(db,col));
            batch.set(r,{...item,id:r.id});
          });
          seedCol(INIT_LEAVES,'leaves');
          seedCol(INIT_CLAIMS,'claims');
          seedCol(INIT_MEMOS,'memos');
          seedCol(INIT_POLICIES,'policies');
          seedCol(INIT_TRAININGS,'trainings');
          seedCol(INIT_FEEDBACKS,'feedbacks');
          seedCol(INIT_FEEDBACK_SECTIONS,'feedbackSections');
          seedCol(INIT_ATTENDANCE,'attendance');
          seedCol(INIT_SHIFTS,'shifts');
          seedCol(INIT_CALENDAR_EVENTS,'calEvents');
          batch.set(doc(collection(db,'projects')),{...INIT_PROJECTS[0],id:""});
          batch.set(doc(db,'company','settings'),{...INIT_COMPANY,id:'settings'});
          await batch.commit();
          console.log('Custera.HR: Database seeded successfully.');
        }
      }catch(e){console.error('Seed error:',e);}
    }
    init();
    const subs=[
      onSnapshot(collection(db,'employees'),s=>{setEmployees(snap2arr(s));setLoading(false);}),
      onSnapshot(collection(db,'users'),s=>setUsers(snap2arr(s))),
      onSnapshot(collection(db,'leaveTypes'),s=>{const d=snap2arr(s);if(d.length)setLeaveTypes(d);}),
      onSnapshot(collection(db,'leaves'),s=>setLeaves(snap2arr(s))),
      onSnapshot(collection(db,'claims'),s=>setClaims(snap2arr(s))),
      onSnapshot(collection(db,'payroll'),s=>setPayroll(snap2arr(s))),
      onSnapshot(collection(db,'memos'),s=>setMemos(snap2arr(s))),
      onSnapshot(collection(db,'policies'),s=>setPolicies(snap2arr(s))),
      onSnapshot(collection(db,'trainings'),s=>setTrainings(snap2arr(s))),
      onSnapshot(collection(db,'feedbacks'),s=>setFeedbacks(snap2arr(s))),
      onSnapshot(collection(db,'feedbackSections'),s=>setFeedbackSections(snap2arr(s))),
      onSnapshot(collection(db,'appraisalForms'),s=>setAppraisalForms(snap2arr(s))),
      onSnapshot(collection(db,'appraisalSubs'),s=>setAppraisalSubs(snap2arr(s))),
      onSnapshot(collection(db,'attendance'),s=>setAttendance(snap2arr(s))),
      onSnapshot(collection(db,'shifts'),s=>setShifts(snap2arr(s))),
      onSnapshot(collection(db,'calEvents'),s=>setCalEvents(snap2arr(s))),
      onSnapshot(collection(db,'projects'),s=>{const d=snap2arr(s);if(d.length)setProjects(d);}),
      onSnapshot(collection(db,'company'),s=>{const d=snap2arr(s);if(d.length)setCompany(d[0]);}),
      onSnapshot(collection(db,'benefitTypes'),s=>{const d=snap2arr(s);if(d.length)setBenefitTypes(d);}),
      onSnapshot(collection(db,'benefitClaims'),s=>setBenefitClaims(snap2arr(s))),
      onSnapshot(collection(db,'allowanceTypes'),s=>{const d=snap2arr(s);if(d.length)setAllowanceTypes(d);}),
      onSnapshot(collection(db,'empFieldConfig'),s=>{const d=snap2arr(s);if(d.length)setEmpFieldConfig(d);}),
    ];
    return()=>subs.forEach(u=>u());
  },[]);

  const isAdmin=session&&(session.role==="admin"||session.role==="superadmin");
  const isSupervisor=session?.role==="supervisor";
  const isAdminOrSup=isAdmin||isSupervisor;
  const empId=session?.empId;
  const NAV=isAdminOrSup?ADMIN_NAV:EMP_NAV;
  const unreadMemos=memos.filter(m=>!(m.readBy||[]).includes(empId));
  const showMemoPopup=!isAdminOrSup&&unreadMemos.length>0;

  // ── Loading screen ────────────────────────────────────────────────────────
  if(loading)return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,flexDirection:"column",gap:16}}>
      <div style={{width:56,height:56,background:C.accent,borderRadius:14,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:800,color:"#fff"}}>C.HR</div>
      <div style={{fontSize:14,color:C.muted,fontWeight:600}}>Custera.HR</div>
      <div style={{fontSize:12,color:C.muted}}>Connecting to database...</div>
      <div style={{width:200,height:4,background:C.border,borderRadius:99,overflow:"hidden"}}>
        <div style={{width:"60%",height:"100%",background:C.accent,borderRadius:99,animation:"none"}}/>
      </div>
    </div>
  );

  if(!session)return <Login users={users} onLogin={u=>{setSession(u);setPage("dashboard");}} lang={lang} setLang={setLang}/>;

  // ── EMPLOYEES ─────────────────────────────────────────────────────────────
  function addEmp(d){
    const id=d.id||"C"+String(employees.length+1).padStart(3,"0");
    fsSet('employees',id,{...d,id});
    toast_("Employee added.");
  }
  function editEmp(id,d){fsUp('employees',id,{...d,id});toast_("Employee updated.");}
  function delEmp(id){fsDel('employees',id);toast_("Employee deleted.","error");}

  // ── USERS ─────────────────────────────────────────────────────────────────
  function addUser(d){const id=d.id||("u"+uid());fsSet('users',id,{...d,id});toast_("Account created.");}
  function editUser(id,d){fsUp('users',id,{...d,id});toast_("Account updated.");}
  function delUser(id){fsDel('users',id);toast_("Account deleted.","error");}
  function resetPw(id,pw){fsUp('users',id,{password:pw});toast_("Password reset.");}

  // ── LEAVES ────────────────────────────────────────────────────────────────
  function addLeave(d){fsAddId('leaves',{...d});toast_("Leave submitted.");}
  function leaveAction(id,status,by,comment){
    const l=leaves.find(x=>x.id===id);if(!l)return;
    const upd={status};
    if(l.status==="Pending Supervisor")upd.supervisorComment=comment||by||"";
    else if(l.status==="Pending HR")upd.hrComment=comment||by||"";
    fsUp('leaves',id,upd);
    toast_(status==="Approved"?"Leave approved":status==="Rejected"?"Leave rejected":"Updated",status==="Rejected"?"error":"success");
  }
  function delLeave(id){fsDel('leaves',id);toast_("Deleted.","error");}
  function addLT(d){fsAdd('leaveTypes',d);}
  function editLT(i,d){if(leaveTypes[i]?.id)fsUp('leaveTypes',leaveTypes[i].id,d);}
  function delLT(i){if(leaveTypes[i]?.id)fsDel('leaveTypes',leaveTypes[i].id);}

  // ── CLAIMS ────────────────────────────────────────────────────────────────
  function addClaim(d){fsAddId('claims',{...d});toast_("Claim submitted.");}
  function claimAction(id,status,by,comment){
    const c=claims.find(x=>x.id===id);if(!c)return;
    const upd={status};
    if(c.status==="Pending Supervisor")upd.supervisorComment=by||"";
    else if(c.status==="Pending HR")upd.hrComment=by||"";
    fsUp('claims',id,upd);toast_(status,"success");
  }
  function delClaim(id){fsDel('claims',id);toast_("Deleted.","error");}

  // ── ATTENDANCE ────────────────────────────────────────────────────────────
  function clockAction(rec){
    const existing=attendance.find(a=>a.empId===rec.empId&&a.date===rec.date);
    if(existing){fsUp('attendance',existing.id,{...rec,id:existing.id});}
    else{fsAddId('attendance',{...rec});}
    toast_("Attendance recorded.");
  }

  // ── PAYROLL ───────────────────────────────────────────────────────────────
  function processPayroll(d){fsAddId('payroll',{...d});toast_("Payroll processed.");}
  function delPayroll(id){fsDel('payroll',id);toast_("Deleted.","error");}
  function publishPayslip(id){fsUp('payroll',id,{status:"Published",publishedOn:todayStr()});toast_("Payslip published to employee.");}

  // ── MEMOS ─────────────────────────────────────────────────────────────────
  function addMemo(d){fsAddId('memos',{...d,readBy:[]});toast_("Announcement posted.");}
  function editMemo(id,d){fsUp('memos',id,d);toast_("Updated.");}
  function delMemo(id){fsDel('memos',id);toast_("Deleted.","error");}
  function readMemo(id){
    const m=memos.find(x=>x.id===id);if(!m)return;
    fsUp('memos',id,{readBy:[...(m.readBy||[]),empId]});
  }

  // ── POLICIES ──────────────────────────────────────────────────────────────
  function addPolicy(d){fsAddId('policies',d);toast_("Policy added.");}
  function editPolicy(id,d){fsUp('policies',id,d);toast_("Updated.");}
  function delPolicy(id){fsDel('policies',id);toast_("Deleted.","error");}

  // ── TRAININGS ─────────────────────────────────────────────────────────────
  function addTraining(d){fsAddId('trainings',d);toast_("Training added.");}
  function editTraining(id,d){fsUp('trainings',id,d);toast_("Updated.");}
  function delTraining(id){fsDel('trainings',id);toast_("Deleted.","error");}

  // ── FEEDBACK ──────────────────────────────────────────────────────────────
  function addFeedback(d){fsAddId('feedbacks',d);toast_("Feedback submitted.");}
  function replyFeedback(id,reply){fsUp('feedbacks',id,{adminReply:reply,status:"Closed"});toast_("Reply sent.");}
  function addFeedbackSection(d){fsAddId('feedbackSections',d);toast_("Category added.");}
  function editFeedbackSection(id,d){fsUp('feedbackSections',id,d);toast_("Updated.");}
  function delFeedbackSection(id){fsDel('feedbackSections',id);toast_("Deleted.","error");}

  // ── APPRAISALS ────────────────────────────────────────────────────────────
  function addAppraisalForm(d){fsAddId('appraisalForms',d);toast_("Form created.");}
  function editAppraisalForm(id,d){fsUp('appraisalForms',id,d);toast_("Updated.");}
  function delAppraisalForm(id){fsDel('appraisalForms',id);toast_("Deleted.","error");}
  function saveAppraisalSub(d){
    const ex=appraisalSubs.find(s=>s.formId===d.formId&&s.empId===d.empId);
    if(ex){fsUp('appraisalSubs',ex.id,{...d,id:ex.id});}
    else{fsAddId('appraisalSubs',d);}
    toast_("Appraisal saved.");
  }

  // ── SHIFTS ────────────────────────────────────────────────────────────────
  function addShift(d){fsAddId('shifts',d);toast_("Shift assigned.");}
  function editShift(id,d){fsUp('shifts',id,d);toast_("Updated.");}
  function delShift(id){fsDel('shifts',id);toast_("Deleted.","error");}

  // ── CALENDAR ──────────────────────────────────────────────────────────────
  function addCalEvent(d){fsAddId('calEvents',d);toast_("Event added.");}
  // Benefits
  function addBenefitType(d){fsSet('benefitTypes',d.id||"bt"+uid(),d);toast_("Benefit type added.");}
  function editBenefitType(id,d){fsUp('benefitTypes',id,d);toast_("Updated.");}
  function addBenefitClaim(d){fsAddId('benefitClaims',{...d});toast_("Claim submitted.");}
  function approveBenefitClaim(id){fsUp('benefitClaims',id,{status:"Approved"});toast_("Claim approved.");}
  function rejectBenefitClaim(id){fsUp('benefitClaims',id,{status:"Rejected"});toast_("Claim rejected.","error");}
  function updateAllowanceTypes(d){d.forEach(a=>fsSet('allowanceTypes',a.id,a));toast_("Allowance types updated.");}
  function updateEmpFieldConfig(d){d.forEach((sec,i)=>fsSet('empFieldConfig',`sec_${i}`,sec));toast_("Field config updated.");}
  function addProject(d){fsAddId('projects',{...d,id:""}).then(()=>toast_("Project added."));}
  function editProject(id,d){fsUp('projects',id,d);toast_("Project updated.");}
  function delProject(id){fsDel('projects',id);toast_("Deleted.","error");}
  function editCalEvent(id,d){fsUp('calEvents',id,d);toast_("Updated.");}
  function delCalEvent(id){fsDel('calEvents',id);toast_("Deleted.","error");}

  // ── PAGE RENDERER ─────────────────────────────────────────────────────────
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
        case "benefits":return <EmpBenefits empId={empId} benefitTypes={benefitTypes} benefitClaims={benefitClaims} onAdd={addBenefitClaim} lang={lang}/>;
        default:return null;
      }
    }
    switch(page){
      case "dashboard":return <AdminDashboard employees={employees} leaves={leaves} claims={claims} attendance={attendance} announcements={memos} calEvents={calEvents} onNavigate={setPage} onLeaveAction={leaveAction} session={session} lang={lang}/>;
      case "employees":return <EmpList employees={employees} leaves={leaves} onAdd={addEmp} onEdit={editEmp} onDelete={delEmp} lang={lang}/>;
      case "users":return <SettingsHub users={users} employees={employees} onAddUser={addUser} onEditUser={editUser} onDeleteUser={delUser} onResetPw={resetPw} empFieldConfig={empFieldConfig} onUpdateFieldConfig={updateEmpFieldConfig} allowanceTypes={allowanceTypes} onUpdateAllowanceTypes={updateAllowanceTypes} lang={lang}/>;
      case "attendance":return <AdminAttendance attendance={attendance} employees={employees} onAdd={clockAction} onEdit={(id,d)=>{const ex=attendance.find(a=>a.id===id);if(ex)clockAction({...ex,...d,id});}} onDelete={id=>{fsDel("attendance",id);toast_("Deleted.","error");}} lang={lang}/>;
      case "leaves":return <AdminLeave leaves={leaves} employees={employees} leaveTypes={leaveTypes} users={users} onAction={leaveAction} onAdd={addLeave} onDelete={delLeave} onAddType={addLT} onEditType={editLT} onDeleteType={delLT} session={session} lang={lang}/>;
      case "claims":return <AdminClaims claims={claims} employees={employees} users={users} session={session} onAction={claimAction} onDelete={delClaim} lang={lang}/>;
      case "payroll":return <AdminPayroll employees={employees} payroll={payroll} claims={claims} onProcess={processPayroll} onDelete={delPayroll} onPublish={publishPayslip} company={company} lang={lang}/>;
      case "memo":return <AdminMemo memos={memos} onAdd={addMemo} onEdit={editMemo} onDelete={delMemo} session={session} lang={lang}/>;
      case "policy":return <AdminPolicy policies={policies} onAdd={addPolicy} onEdit={editPolicy} onDelete={delPolicy} session={session} lang={lang}/>;
      case "training":return <AdminTraining trainings={trainings} employees={employees} onAdd={addTraining} onEdit={editTraining} onDelete={delTraining} session={session} lang={lang}/>;
      case "feedback":return <AdminFeedback feedbacks={feedbacks} sections={feedbackSections} employees={employees} onReply={replyFeedback} onDeleteFeedback={id=>fsDel('feedbacks',id)} onAddSection={addFeedbackSection} onEditSection={editFeedbackSection} onDeleteSection={delFeedbackSection} lang={lang}/>;
      case "appraisal":return <AdminAppraisal forms={appraisalForms} submissions={appraisalSubs} employees={employees} onAddForm={addAppraisalForm} onEditForm={editAppraisalForm} onDeleteForm={delAppraisalForm} onSaveSubmission={saveAppraisalSub} session={session} lang={lang}/>;
      case "calendar":return <AdminCalendar events={calEvents} leaves={leaves} employees={employees} onAdd={addCalEvent} onEdit={editCalEvent} onDelete={delCalEvent} session={session} lang={lang}/>;
      case "shift":return <AdminShift shifts={shifts} employees={employees} onAdd={addShift} onEdit={editShift} onDelete={delShift} lang={lang}/>
      case "benefits":return <AdminBenefits benefitTypes={benefitTypes} benefitClaims={benefitClaims} employees={employees} onAddType={addBenefitType} onEditType={editBenefitType} onDeleteType={id=>fsDel("benefitTypes",id)} onApprove={approveBenefitClaim} onReject={rejectBenefitClaim} lang={lang}/>
      case "projects":return <AdminProjects projects={projects} employees={employees} onAdd={addProject} onEdit={editProject} onDelete={delProject} lang={lang}/>
      case "tracker":return <AdminTracker employees={employees} leaves={leaves} leaveTypes={leaveTypes} lang={lang}/>;
      default:return null;
    }
  }

  // ── LAYOUT ────────────────────────────────────────────────────────────────
  return(
    <ErrorBoundary>
    <div style={{display:"flex",height:"100vh",fontFamily:"'Inter',-apple-system,BlinkMacSystemFont,sans-serif",background:C.bg,overflow:"hidden"}}>
      <aside style={{width:190,background:C.sidebar,display:"flex",flexDirection:"column",flexShrink:0,overflowY:"auto"}}>
        <div style={{padding:"16px 14px 12px",borderBottom:"1px solid rgba(255,255,255,0.08)"}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:34,height:34,background:`linear-gradient(135deg,${C.accent},${C.purple})`,borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"#fff",flexShrink:0,boxShadow:"0 4px 12px rgba(99,102,241,0.4)"}}>C·HR</div>
            <div><div style={{color:"#fff",fontWeight:800,fontSize:13,lineHeight:1.1}}>Custera<span style={{color:"#60A5FA"}}>.HR</span></div><div style={{color:"rgba(255,255,255,0.4)",fontSize:9,marginTop:1}}>SG Construction</div></div>
          </div>
          <div style={{display:"flex",gap:4,marginTop:10}}>
            {["en","zh"].map(l=><button key={l} onClick={()=>setLang(l)} style={{flex:1,background:lang===l?C.accent:"rgba(255,255,255,0.08)",color:lang===l?"#fff":"rgba(255,255,255,0.5)",border:"none",borderRadius:5,padding:"3px 0",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{l==="en"?"EN":"中文"}</button>)}
          </div>
        </div>
        <nav style={{flex:1,padding:"8px 6px"}}>
          {NAV.map(n=>{const active=page===n.id;return(
            <button key={n.id} onClick={()=>setPage(n.id)}
              style={{display:"flex",alignItems:"center",gap:8,width:"100%",padding:"8px 9px",borderRadius:7,border:"none",background:active?C.sidebarA:"transparent",color:active?"#93C5FD":C.sidebarT,fontWeight:active?700:500,fontSize:12,cursor:"pointer",marginBottom:1,textAlign:"left",fontFamily:"inherit"}}
              onMouseEnter={e=>{if(!active)e.currentTarget.style.background="rgba(255,255,255,0.07)";}}
              onMouseLeave={e=>{if(!active)e.currentTarget.style.background="transparent";}}>
              <span style={{width:18,height:18,background:"rgba(255,255,255,0.1)",borderRadius:4,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:800,flexShrink:0}}>{n.icon}</span>
              {t(lang,n.l)}
              {n.id==="memo"&&!isAdminOrSup&&unreadMemos.length>0&&<span style={{marginLeft:"auto",background:C.danger,color:"#fff",fontSize:9,fontWeight:800,borderRadius:99,padding:"1px 5px",flexShrink:0}}>{unreadMemos.length}</span>}
            </button>
          );})}
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
    </div></ErrorBoundary>
  );
}


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
        {[["Employee ID","id","text"],["Full Name","name","text"],["FIN/NRIC","fin","text"],["Department","department","text"],["Company","company","text"],["Position","position","text"],["Date Joined","dateJoined","date"],["Date of Birth","dob","date"],["Nationality","nationality","text"],["Work Pass","workPass","text"],["Mobile","mobile","text"],["Basic Salary","basicSalary","number"],["Allowance","allowance","number"],["Status","status","text"],["Work Pass Issue Date","workPassIssueDate","date"],["Work Pass Expiry Date","workPassExpiryDate","date"],["Probation End Date","probationEndDate","date"],["Passport Number","passportNumber","text"],["Passport Issue Date","passportIssueDate","date"],["Passport Expiry Date","passportExpiryDate","date"],["Qualification","qualification","text"],["Phone Allowance (S$)","phoneAllowance","number"],["Housing Allowance (S$)","housingAllowance","number"],["Airfare Allowance (S$)","airfareAllowance","number"],["Other Allowance (S$)","otherAllowance","number"]].map(([l,k,tp])=><Inp key={k} label={l} type={tp} value={f[k]||""} onChange={v=>setFk(k,v)} readOnly={k==="id"&&!!editE}/>)}
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
