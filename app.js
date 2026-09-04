const STORES={students:"umi_students_v3",payments:"umi_payments_v3",discounts:"umi_discounts_v4",departments:"umi_departments_v3",subjects:"umi_subjects_v3",grades:"umi_grades_v3"};
const defaults={
students:[
{id:"s1",number:"10001",name:"أحمد محمد علي",dept:"علوم مالية ومصرفية",stage:"الأولى",study:"صباحية",fee:1800000,paid:900000},
{id:"s2",number:"10002",name:"سارة حسن كريم",dept:"إدارة الأعمال",stage:"الثانية",study:"مسائية",fee:1600000,paid:1200000},
{id:"s3",number:"10003",name:"علي ماجد جواد",dept:"القانون",stage:"الثالثة",study:"صباحية",fee:2000000,paid:2000000}],
departments:[
{id:"d1",name:"علوم مالية ومصرفية",code:"FIN"},
{id:"d2",name:"إدارة الأعمال",code:"BUS"},
{id:"d3",name:"القانون",code:"LAW"}],
subjects:[
{id:"m1",name:"مبادئ المحاسبة",dept:"علوم مالية ومصرفية",stage:"الأولى",units:3},
{id:"m2",name:"مبادئ الإدارة",dept:"إدارة الأعمال",stage:"الثانية",units:3},
{id:"m3",name:"القانون المدني",dept:"القانون",stage:"الثالثة",units:4}]
};
const load=(key,def)=>JSON.parse(localStorage.getItem(STORES[key])||"null")??def;
let students=load("students",defaults.students),payments=load("payments",[]),discounts=load("discounts",[]),departments=load("departments",defaults.departments),subjects=load("subjects",defaults.subjects),grades=load("grades",[]);
const qs=s=>document.querySelector(s), qsa=s=>document.querySelectorAll(s);
const normalizeDigits=v=>String(v??"").replace(/[٠-٩۰-۹]/g,c=>String("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹".indexOf(c)%10));
const num=v=>Number(normalizeDigits(v).replace(/,/g,"").trim()||0);
const fmt=n=>new Intl.NumberFormat("en-US").format(Number(n||0))+" د.ع";
const save=()=>{localStorage.setItem(STORES.students,JSON.stringify(students));localStorage.setItem(STORES.payments,JSON.stringify(payments));localStorage.setItem(STORES.discounts,JSON.stringify(discounts));localStorage.setItem(STORES.departments,JSON.stringify(departments));localStorage.setItem(STORES.subjects,JSON.stringify(subjects));localStorage.setItem(STORES.grades,JSON.stringify(grades));
qs("#financeSearch")?.addEventListener("input",renderFinance);qs("#financeStatusFilter")?.addEventListener("change",renderFinance);qs("#financeDeptFilter")?.addEventListener("change",renderFinance);
function printReceipt(payment){
const s=students.find(x=>x.id===payment.studentId),disc=getDiscount(s.id),paid=getPaid(s.id),due=Math.max(0,num(s.fee)-disc),rem=Math.max(0,due-paid);
const w=window.open("","_blank","width=820,height=700");if(!w)return;
w.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><title>إيصال دفع</title><style>body{font-family:Tahoma,Arial;padding:30px;color:#17243a}.head{display:flex;justify-content:space-between;border-bottom:2px solid #17243a;padding-bottom:14px}.box{border:1px solid #ddd;padding:14px;border-radius:8px;margin-top:14px}table{width:100%;border-collapse:collapse;margin-top:16px}td,th{border:1px solid #ddd;padding:10px;text-align:right}.total{font-size:18px;font-weight:bold}</style></head><body><div class="head"><div><h2>UniManager IQ</h2><div>إيصال دفع مالي</div></div><div>التاريخ: ${payment.date}</div></div><div class="box"><p><b>الطالب:</b> ${s.name}</p><p><b>الرقم:</b> ${s.number}</p><p><b>القسم:</b> ${s.dept}</p><p><b>المرحلة:</b> ${s.stage}</p></div><table><tr><th>البيان</th><th>القيمة</th></tr><tr><td>القسط</td><td>${fmt(s.fee)}</td></tr><tr><td>الخصومات</td><td>${fmt(disc)}</td></tr><tr><td>المبلغ المستحق</td><td>${fmt(due)}</td></tr><tr><td>هذه الدفعة</td><td class="total">${fmt(payment.amount)}</td></tr><tr><td>إجمالي المدفوع</td><td>${fmt(paid)}</td></tr><tr><td>المتبقي</td><td>${fmt(rem)}</td></tr></table><div class="box">طريقة الدفع: ${payment.method} ${payment.installment?` — الدفعة رقم ${payment.installment}`:""}<br>${payment.note||""}</div><p style="margin-top:45px">توقيع الموظف: ____________________</p><script>window.print()<\/script></body></html>`);w.document.close();}


qs("#downloadSubjectsTemplateBtn").onclick=()=>{
download([{القسم:departments[0]?.name||"اسم القسم",المرحلة:"الأولى",المادة:"اسم المادة",الوحدات:3},{القسم:departments[0]?.name||"اسم القسم",المرحلة:"الثانية",المادة:"اسم المادة 2",الوحدات:3}],"subjects-import-template.xlsx");
};
qs("#importSubjectsBtn").onclick=()=>qs("#subjectsExcelInput").click();
qs("#subjectsExcelInput").onchange=async e=>{
const f=e.target.files[0];if(!f)return;
try{
const wb=XLSX.read(await f.arrayBuffer());
const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});
let added=0,skipped=0;
for(const r of rows){
const dept=String(r["القسم"]||r["اسم القسم"]||r.dept||"").trim();
const stage=String(r["المرحلة"]||r.stage||"").trim();
const name=String(r["المادة"]||r["اسم المادة"]||r.name||"").trim();
const units=num(r["الوحدات"]||r["وحدات المادة"]||r.units||0);
if(!dept||!stage||!name||!units||!departments.some(d=>d.name===dept)||!["الأولى","الثانية","الثالثة","الرابعة"].includes(stage)){skipped++;continue;}
const duplicate=subjects.some(m=>m.dept===dept&&m.stage===stage&&m.name===name);
if(duplicate){skipped++;continue;}
subjects.push({id:crypto.randomUUID(),name,dept,stage,units});added++;
}
if(!added&&skipped)throw new Error("لم تتم إضافة مواد. تأكد من أسماء الأقسام والمراحل والأعمدة في القالب.");
save();alert(`تمت إضافة ${added} مادة${skipped?` وتجاوز ${skipped} صف`:""}.`);
}catch(err){alert("تعذر استيراد مواد الأقسام: "+err.message)}
e.target.value="";
};
renderAll();};
function showPage(page){qsa(".page").forEach(x=>x.classList.add("hidden"));qs("#"+page+"Page").classList.remove("hidden");qsa(".nav-item").forEach(x=>x.classList.toggle("active",x.dataset.page===page));qs("#pageTitle").textContent={dashboard:"لوحة التحكم",students:"إدارة الطلبة",finance:"الأقساط والحسابات",grades:"الدرجات و Master Sheet",departments:"الأقسام والمواد",reports:"التقارير"}[page]||page;}
qsa(".nav-item").forEach(b=>b.onclick=()=>showPage(b.dataset.page));qsa("[data-page-jump]").forEach(b=>b.onclick=()=>showPage(b.dataset.pageJump));
qs("#loginForm").onsubmit=e=>{e.preventDefault();if(qs("#username").value==="admin"&&qs("#password").value==="admin123"){qs("#loginView").classList.add("hidden");qs("#appView").classList.remove("hidden");}else alert("بيانات الدخول غير صحيحة");};
qs("#logoutBtn").onclick=()=>{qs("#appView").classList.add("hidden");qs("#loginView").classList.remove("hidden");};
qs("#dateNow").textContent=new Intl.DateTimeFormat("ar-IQ",{dateStyle:"full"}).format(new Date());
function statusBadge(s){const r=Math.max(0,s.fee-s.paid),p=s.fee?Math.round(s.paid/s.fee*100):0;if(r<=0)return '<span class="badge success">مكتمل</span>';if(p>=50)return '<span class="badge warning">جزئي</span>';return '<span class="badge danger">متأخر</span>';}
function renderDashboard(){const exp=students.reduce((a,s)=>a+num(s.fee),0),paid=students.reduce((a,s)=>a+num(s.paid),0),rem=exp-paid,pct=exp?Math.round(paid/exp*100):0;qs("#statStudents").textContent=students.length;qs("#studentsHeroCount")&&(qs("#studentsHeroCount").textContent=students.length);qs("#financeHeroRemaining")&&(qs("#financeHeroRemaining").textContent=fmt(rem));qs("#gradesHeroCount")&&(qs("#gradesHeroCount").textContent=grades.length);qs("#deptHeroCount")&&(qs("#deptHeroCount").textContent=departments.length);qs("#statExpected").textContent=fmt(exp);qs("#statPaid").textContent=fmt(paid);qs("#statRemaining").textContent=fmt(rem);qs("#collectionPct").textContent=pct+"%";qs("#recentStudents").innerHTML=students.slice(-5).reverse().map(s=>`<div class="recent-row"><span>${s.name}<small class="muted"> — ${s.dept}</small></span><strong>${fmt(s.paid)}</strong></div>`).join("")||'<div class="empty-state">لا توجد بيانات</div>';}
function fillDeptSelect(id,all=true){const el=qs("#"+id),cur=el.value;el.innerHTML=(all?'<option value="">كل الأقسام</option>':"<option value=\"\">اختر القسم</option>")+departments.map(d=>`<option value="${d.name}">${d.name}</option>`).join("");el.value=departments.some(d=>d.name===cur)?cur:"";}
function renderStudentFilters(){fillDeptSelect("studentDeptFilter");fillDeptSelect("gradeDeptFilter");fillDeptSelect("sDept",false);fillDeptSelect("subDept",false);}
function renderStudentMiniStats(){
const total=students.length;
const morning=students.filter(s=>s.study==="صباحية").length;
const evening=students.filter(s=>s.study==="مسائية").length;
const due=students.filter(s=>Math.max(0,num(s.fee)-num(s.paid))>0).length;
if(qs("#studentTotalMini"))qs("#studentTotalMini").textContent=total;
if(qs("#studentMorningMini"))qs("#studentMorningMini").textContent=morning;
if(qs("#studentEveningMini"))qs("#studentEveningMini").textContent=evening;
if(qs("#studentDueMini"))qs("#studentDueMini").textContent=due;
}
function renderStudents(){renderStudentMiniStats();renderStudentFilters();const q=qs("#studentSearch").value.trim().toLowerCase(),d=qs("#studentDeptFilter").value,st=qs("#studentStageFilter").value;const rows=students.filter(s=>(!q||[s.number,s.name,s.dept,s.stage].join(" ").toLowerCase().includes(q))&&(!d||s.dept===d)&&(!st||s.stage===st));qs("#studentsTable").innerHTML=rows.map(s=>`<tr><td>${s.number}</td><td><strong>${s.name}</strong></td><td>${s.dept}</td><td>${s.stage}</td><td>${s.study}</td><td>${fmt(s.fee)}</td><td>${fmt(s.paid)}</td><td>${statusBadge(s)}</td><td><button class="action-btn" onclick="openProfile('${s.id}')">ملف</button><button class="action-btn" onclick="editStudent('${s.id}')">تعديل</button><button class="action-btn danger-btn" onclick="deleteStudent('${s.id}')">حذف</button></td></tr>`).join("")||'<tr><td colspan="9" class="empty-state">لا توجد نتائج</td></tr>';}
function getDiscount(studentId){return discounts.filter(d=>d.studentId===studentId).reduce((a,d)=>a+num(d.amount),0);}
function getPaid(studentId){return payments.filter(p=>p.studentId===studentId).reduce((a,p)=>a+num(p.amount),0);}
function renderFinance(){
const exp=students.reduce((a,s)=>a+num(s.fee),0), paid=payments.reduce((a,p)=>a+num(p.amount),0), disc=discounts.reduce((a,d)=>a+num(d.amount),0), remaining=Math.max(0,exp-disc-paid);
if(qs("#financeExpectedKpi"))qs("#financeExpectedKpi").textContent=fmt(exp);
if(qs("#financePaidKpi"))qs("#financePaidKpi").textContent=fmt(paid);
if(qs("#financeDiscountKpi"))qs("#financeDiscountKpi").textContent=fmt(disc);
if(qs("#financeRemainingKpi"))qs("#financeRemainingKpi").textContent=fmt(remaining);
const q=(qs("#financeSearch")?.value||"").trim().toLowerCase(),status=qs("#financeStatusFilter")?.value||"",dept=qs("#financeDeptFilter")?.value||"";
if(qs("#financeDeptFilter")){const cur=qs("#financeDeptFilter").value;qs("#financeDeptFilter").innerHTML='<option value="">كل الأقسام</option>'+departments.map(d=>`<option>${d.name}</option>`).join("");qs("#financeDeptFilter").value=departments.some(d=>d.name===cur)?cur:"";}
const rows=students.filter(s=>(!q||[s.number,s.name].join(" ").toLowerCase().includes(q))&&(!dept||s.dept===dept)).map(s=>{const discount=getDiscount(s.id),paidNow=getPaid(s.id),due=Math.max(0,num(s.fee)-discount),rem=Math.max(0,due-paidNow),pct=due?Math.round(paidNow/due*100):100;const st=rem<=0?"مكتمل":paidNow>0?"جزئي":"متأخر";return {s,discount,due,paidNow,rem,pct,st};}).filter(x=>!status||x.st===status);
qs("#financeTable").innerHTML=rows.map(x=>`<tr><td><strong>${x.s.name}</strong><div class="muted">${x.s.number}</div></td><td>${fmt(x.s.fee)}</td><td>${fmt(x.discount)}</td><td>${fmt(x.due)}</td><td>${fmt(x.paidNow)}</td><td class="finance-remain">${fmt(x.rem)}</td><td><span class="badge ${x.st==="مكتمل"?"success":x.st==="جزئي"?"warning":"danger"}">${x.st}</span></td><td><button class="action-btn" onclick="openPaymentFor('${x.s.id}')">دفعة</button><button class="action-btn" onclick="openDiscountFor('${x.s.id}')">خصم</button><button class="action-btn" onclick="openProfile('${x.s.id}')">كشف</button></td></tr>`).join("")||'<tr><td colspan="8" class="empty-state">لا توجد نتائج</td></tr>';
}

const DEFAULT_DEPT_LOGOS = {
  FIN: "data:image/svg+xml;charset=UTF-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'%3E%3Crect width='120' height='120' rx='24' fill='%23eaf2ff'/%3E%3Ccircle cx='60' cy='60' r='40' fill='%231d4ed8'/%3E%3Ctext x='60' y='73' font-size='42' text-anchor='middle' fill='white' font-family='Arial'%3Eد%3C/text%3E%3C/svg%3E",
  BUS: "data:image/svg+xml;charset=UTF-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'%3E%3Crect width='120' height='120' rx='24' fill='%23eefbf3'/%3E%3Ccircle cx='60' cy='60' r='40' fill='%2315803d'/%3E%3Ctext x='60' y='73' font-size='42' text-anchor='middle' fill='white' font-family='Arial'%3Eأ%3C/text%3E%3C/svg%3E",
  LAW: "data:image/svg+xml;charset=UTF-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'%3E%3Crect width='120' height='120' rx='24' fill='%23fff7ed'/%3E%3Ccircle cx='60' cy='60' r='40' fill='%23b45309'/%3E%3Ctext x='60' y='73' font-size='42' text-anchor='middle' fill='white' font-family='Arial'%3Eق%3C/text%3E%3C/svg%3E"
};
function getDeptLogo(dept){
  return dept.logo || DEFAULT_DEPT_LOGOS[dept.code] || DEFAULT_DEPT_LOGOS.LAW;
}
function readImageAsDataURL(file){
  return new Promise((resolve,reject)=>{
    if(!file){resolve("");return;}
    const reader=new FileReader();
    reader.onload=()=>resolve(reader.result);
    reader.onerror=reject;
    reader.readAsDataURL(file);
  });
}
window.changeDeptLogo=async function(id){
  const dept=departments.find(d=>d.id===id);
  if(!dept)return;
  const input=document.createElement("input");
  input.type="file";
  input.accept="image/*";
  input.onchange=async()=>{
    if(!input.files?.[0])return;
    dept.logo=await readImageAsDataURL(input.files[0]);
    save();
    renderDepartments();
  };
  input.click();
};

function renderDepartments(){
if(qs("#departmentsList")){
qs("#departmentsList").innerHTML=departments.map(d=>{
const deptSubjects=subjects.filter(m=>m.dept===d.name);
const stages=["الأولى","الثانية","الثالثة","الرابعة"];
const stagesHtml=stages.map(stage=>{
const list=deptSubjects.filter(m=>m.stage===stage);
return `<div class="department-stage"><div class="stage-title">${stage} — ${list.length} مادة</div>${list.length?list.map(m=>`<span class="subject-chip">${m.name}<small> (${m.units} و)</small></span>`).join(""):'<span class="no-subjects">لا توجد مواد لهذه المرحلة.</span>'}</div>`;
}).join("");
return `<div class="department-card"><div class="department-card-head"><div class="department-identity"><img class="department-logo" src="${getDeptLogo(d)}" alt="شعار ${d.name}"><div><div class="department-title">${d.name}</div><div class="department-code">رمز: ${d.code}</div></div></div><div class="list-meta"><span class="subject-count">${deptSubjects.length}</span><button class="small-btn" onclick="changeDeptLogo('${d.id}')">تغيير الشعار</button><button class="small-btn" onclick="deleteDept('${d.id}')">حذف</button></div></div>${stagesHtml}</div>`;
}).join("")||'<div class="empty-state">لا توجد أقسام.</div>';
}
if(qs("#subjectsList")){
qs("#subjectsList").innerHTML=`<div class="subject-management">${departments.map(d=>{
const list=subjects.filter(m=>m.dept===d.name).sort((a,b)=>["الأولى","الثانية","الثالثة","الرابعة"].indexOf(a.stage)-["الأولى","الثانية","الثالثة","الرابعة"].indexOf(b.stage)||a.name.localeCompare(b.name,"ar"));
return `<div class="subject-management-card"><div class="subject-management-head"><div><div class="subject-management-title">${d.name}</div><div class="subject-management-meta">مجموعة المواد الخاصة بالقسم</div></div><span class="subject-count">${list.length} مادة</span></div><div class="subjects-table-wrap"><table class="subjects-table"><thead><tr><th>المادة</th><th>المرحلة</th><th>الوحدات</th><th>إجراء</th></tr></thead><tbody>${list.length?list.map(m=>`<tr><td><strong>${m.name}</strong></td><td><span class="subject-stage-chip">${m.stage}</span></td><td>${m.units}</td><td><button class="action-btn danger-btn" onclick="deleteSubject('${m.id}')">حذف</button></td></tr>`).join(""):`<tr><td colspan="4" class="empty-state">لا توجد مواد مضافة لهذا القسم.</td></tr>`}</tbody></table></div></div>`;
}).join("")||'<div class="empty-state">أضف قسمًا أولًا.</div>'}</div>`;
}
}
function renderGradeFilters(){const sub=qs("#gradeSubjectFilter"),cur=sub.value;sub.innerHTML='<option value="">كل المواد</option>'+subjects.map(s=>`<option value="${s.id}">${s.name}</option>`).join("");sub.value=subjects.some(s=>s.id===cur)?cur:"";}
function renderGrades(){
renderStudentFilters();renderGradeFilters();
const d=qs("#gradeDeptFilter").value,st=qs("#gradeStageFilter").value,sub=qs("#gradeSubjectFilter").value,term=qs("#gradeTermFilter").value;
const rows=grades.filter(g=>{const s=students.find(x=>x.id===g.studentId),m=subjects.find(x=>x.id===g.subjectId);return s&&m&&(!d||s.dept===d)&&(!st||s.stage===st)&&(!sub||g.subjectId===sub)&&(!term||g.term===term);});
const totals=rows.map(g=>Math.min(100,num(g.course)+num(g.final)));
const passCount=totals.filter(x=>x>=50).length,failCount=totals.length-passCount,avg=totals.length?Math.round(totals.reduce((a,b)=>a+b,0)/totals.length):0;
if(qs("#gradesTotalKpi"))qs("#gradesTotalKpi").textContent=rows.length;
if(qs("#gradesPassKpi"))qs("#gradesPassKpi").textContent=passCount;
if(qs("#gradesFailKpi"))qs("#gradesFailKpi").textContent=failCount;
if(qs("#gradesAvgKpi"))qs("#gradesAvgKpi").textContent=avg;
qs("#gradesTable").innerHTML=rows.map(g=>{const s=students.find(x=>x.id===g.studentId),m=subjects.find(x=>x.id===g.subjectId),total=Math.min(100,num(g.course)+num(g.final)),pass=total>=50;return `<tr><td><strong>${s?.name||""}</strong><div class="muted">${s?.number||""}</div></td><td>${m?.name||""}</td><td>${g.term||"الأول"}</td><td>${normalizeDigits(g.course)}</td><td>${normalizeDigits(g.final)}</td><td><strong>${total}</strong></td><td class="${pass?"result-pass":"result-fail"}">${pass?"ناجح":"راسب"}</td><td><button class="action-btn danger-btn" onclick="deleteGrade('${g.id}')">حذف</button></td></tr>`}).join("")||'<tr><td colspan="8" class="empty-state">لا توجد درجات</td></tr>';
}
function renderAll(){renderDashboard();renderStudents();renderFinance();renderDepartments();renderGrades();fillPaymentStudents();}
qs("#studentSearch").oninput=renderStudents;qs("#studentDeptFilter").onchange=renderStudents;qs("#studentStageFilter").onchange=renderStudents;qs("#gradeDeptFilter").onchange=renderGrades;qs("#gradeStageFilter").onchange=renderGrades;qs("#gradeSubjectFilter").onchange=renderGrades;qs("#gradeTermFilter").onchange=renderGrades;
function openStudent(id=""){const s=students.find(x=>x.id===id);qs("#studentDialogTitle").textContent=s?"تعديل طالب":"إضافة طالب";qs("#studentId").value=s?.id||"";qs("#sNumber").value=s?.number||"";qs("#sName").value=s?.name||"";qs("#sDept").value=s?.dept||departments[0]?.name||"";qs("#sStage").value=s?.stage||"الأولى";qs("#sStudy").value=s?.study||"صباحية";qs("#sFee").value=s?.fee??"";qs("#sPaid").value=s?.paid??"";qs("#studentDialog").showModal();}
window.editStudent=openStudent;
window.deleteStudent=id=>{const s=students.find(x=>x.id===id);if(confirm(`حذف الطالب ${s?.name}؟`)){students=students.filter(x=>x.id!==id);payments=payments.filter(p=>p.studentId!==id);grades=grades.filter(g=>g.studentId!==id);save();}};
window.openProfile=id=>{const s=students.find(x=>x.id===id);const rem=Math.max(0,num(s.fee)-num(s.paid));const ps=payments.filter(p=>p.studentId===id).reverse();qs("#studentProfile").innerHTML=`<div class="profile-grid"><p><b>الرقم:</b> ${s.number}</p><p><b>الاسم:</b> ${s.name}</p><p><b>القسم:</b> ${s.dept}</p><p><b>المرحلة:</b> ${s.stage}</p><p><b>الدراسة:</b> ${s.study}</p><p><b>القسط:</b> ${fmt(s.fee)}</p><p><b>المدفوع:</b> ${fmt(s.paid)}</p><p><b>المتبقي:</b> ${fmt(rem)}</p></div><hr><h4>سجل الدفعات</h4>${ps.map(p=>`<div class="list-row"><span>${new Date(p.date).toLocaleDateString("en-GB")} — ${p.method} ${p.note?`— ${p.note}`:""}</span><strong>${fmt(p.amount)}</strong></div>`).join("")||'<div class="empty-state">لا توجد دفعات مسجلة</div>'}`;qs("#studentProfileDialog").showModal();};
window.closeDialog=id=>qs("#"+id).close();
qs("#addStudentBtn").onclick=()=>openStudent();
qs("#studentForm").onsubmit=e=>{e.preventDefault();const id=qs("#studentId").value,data={id:id||crypto.randomUUID(),number:normalizeDigits(qs("#sNumber").value),name:qs("#sName").value.trim(),dept:qs("#sDept").value,stage:qs("#sStage").value,study:qs("#sStudy").value,fee:num(qs("#sFee").value),paid:num(qs("#sPaid").value)};if(id)students=students.map(x=>x.id===id?data:x);else students.push(data);qs("#studentDialog").close();save();};
function fillPaymentStudents(){qs("#pStudent").innerHTML=students.map(s=>`<option value="${s.id}">${s.number} — ${s.name}</option>`).join("");}
function fillDiscountStudents(){qs("#dStudent").innerHTML=students.map(s=>`<option value="${s.id}">${s.number} — ${s.name}</option>`).join("");}
window.openPaymentFor=id=>{fillPaymentStudents();qs("#pStudent").value=id;qs("#pAmount").value="";qs("#pInstallment").value="";qs("#pMethod").value="نقدًا";qs("#pDate").value=new Date().toISOString().slice(0,10);qs("#pNote").value="";qs("#paymentDialog").showModal();};
qs("#addPaymentBtn").onclick=()=>window.openPaymentFor(students[0]?.id||"");
qs("#paymentForm").onsubmit=e=>{e.preventDefault();const studentId=qs("#pStudent").value,amount=num(qs("#pAmount").value),s=students.find(x=>x.id===studentId);if(!s||amount<=0)return;const rem=Math.max(0,num(s.fee)-getDiscount(s.id)-getPaid(s.id));if(amount>rem&&rem>0&&!confirm("المبلغ أكبر من المتبقي. هل تريد المتابعة؟"))return;const payment={id:crypto.randomUUID(),studentId,amount,date:qs("#pDate").value||new Date().toISOString().slice(0,10),method:qs("#pMethod").value,note:qs("#pNote").value.trim(),installment:normalizeDigits(qs("#pInstallment").value)};payments.push(payment);qs("#paymentDialog").close();save();printReceipt(payment);};
window.openDiscountFor=id=>{fillDiscountStudents();qs("#dStudent").value=id;qs("#dAmount").value="";qs("#dReason").value="";qs("#discountDialog").showModal();};
qs("#discountForm").onsubmit=e=>{e.preventDefault();const studentId=qs("#dStudent").value,amount=num(qs("#dAmount").value),reason=qs("#dReason").value.trim();if(amount<=0||!reason)return;discounts.push({id:crypto.randomUUID(),studentId,amount,reason,date:new Date().toISOString().slice(0,10)});qs("#discountDialog").close();save();};
qs("#addGradeBtn").onclick=()=>{
qs("#gStudent").innerHTML=students.map(s=>`<option value="${s.id}" data-dept="${s.dept}" data-stage="${s.stage}">${s.number} — ${s.name}</option>`).join("");
qs("#gDept").innerHTML=departments.map(d=>`<option value="${d.name}">${d.name}</option>`).join("");
qs("#gStage").value="الأولى";
updateGradeEntrySubjects();
qs("#gTerm").value="الأول";qs("#gCourse").value="";qs("#gFinal").value="";qs("#gradeDialog").showModal();
};
function updateGradeEntrySubjects(){
const dept=qs("#gDept").value,stage=qs("#gStage").value;
const list=subjects.filter(m=>m.dept===dept&&m.stage===stage);
qs("#gSubject").innerHTML=list.length?list.map(m=>`<option value="${m.id}">${m.name} — ${m.units} وحدات</option>`).join(""):'<option value="">لا توجد مواد لهذا القسم والمرحلة</option>';
}
qs("#gDept").onchange=()=>{updateGradeEntrySubjects();const s=students.find(x=>x.dept===qs("#gDept").value&&x.stage===qs("#gStage").value);if(s)qs("#gStudent").value=s.id;};
qs("#gStage").onchange=()=>{updateGradeEntrySubjects();const s=students.find(x=>x.dept===qs("#gDept").value&&x.stage===qs("#gStage").value);if(s)qs("#gStudent").value=s.id;};
qs("#gradeForm").onsubmit=e=>{e.preventDefault();const studentId=qs("#gStudent").value,subjectId=qs("#gSubject").value,course=num(qs("#gCourse").value),finalScore=num(qs("#gFinal").value),student=students.find(x=>x.id===studentId),subject=subjects.find(x=>x.id===subjectId);if(!student||!subject||student.dept!==subject.dept||student.stage!==subject.stage){alert("المادة المختارة لا تنتمي إلى قسم الطالب ومرحلته.");return;}if(course<0||course>40||finalScore<0||finalScore>60){alert("السعي يجب أن يكون بين 0 و40 والنهائي بين 0 و60.");return;}const term=qs("#gTerm").value;const old=grades.find(g=>g.studentId===studentId&&g.subjectId===subjectId&&((g.term||"الأول")===term));const data={id:old?.id||crypto.randomUUID(),studentId,subjectId,term,course,final:finalScore};grades=old?grades.map(g=>g.id===old.id?data:g):[...grades,data];qs("#gradeDialog").close();save();};
window.deleteGrade=id=>{grades=grades.filter(g=>g.id!==id);save();};
qs("#addDeptBtn").onclick=()=>{qs("#dName").value="";qs("#dCode").value="";qs("#deptDialog").showModal();};
qs("#deptForm").onsubmit=async e=>{e.preventDefault();const name=qs("#dName").value.trim(),code=normalizeDigits(qs("#dCode").value.trim()).toUpperCase(),logo=await readImageAsDataURL(qs("#dLogo")?.files?.[0]);if(!name||!code){alert("أدخل اسم القسم والرمز.");return;}if(departments.some(d=>d.name===name||d.code===code)){alert("القسم أو رمز القسم موجود مسبقًا.");return;}departments.push({id:crypto.randomUUID(),name,code,logo:logo||DEFAULT_DEPT_LOGOS[code]||DEFAULT_DEPT_LOGOS.LAW});qs("#deptDialog").close();save();};
window.deleteDept=id=>{const d=departments.find(x=>x.id===id);if(students.some(s=>s.dept===d?.name)||subjects.some(s=>s.dept===d?.name)){alert("لا يمكن حذف القسم لأنه مرتبط بطلبة أو مواد.");return;}departments=departments.filter(x=>x.id!==id);save();};
qs("#addSubjectBtn").onclick=()=>{fillDeptSelect("subDept",false);qs("#subName").value="";qs("#subStage").value="الأولى";qs("#subUnits").value="3";qs("#subjectDialog").showModal();};
qs("#subjectForm").onsubmit=e=>{e.preventDefault();subjects.push({id:crypto.randomUUID(),name:qs("#subName").value.trim(),dept:qs("#subDept").value,stage:qs("#subStage").value,units:num(qs("#subUnits").value)});qs("#subjectDialog").close();save();};
window.deleteSubject=id=>{if(grades.some(g=>g.subjectId===id)){alert("لا يمكن حذف المادة لأنها مرتبطة بدرجات.");return;}subjects=subjects.filter(x=>x.id!==id);save();};
function download(rows,name){const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(rows);XLSX.utils.book_append_sheet(wb,ws,"البيانات");XLSX.writeFile(wb,name);}
function exportStudents(){download(students.map(s=>({الرقم:s.number,"اسم الطالب":s.name,القسم:s.dept,المرحلة:s.stage,الدراسة:s.study,"القسط الكلي":s.fee,"المدفوع":s.paid,"المتبقي":Math.max(0,s.fee-s.paid)})),"students-report.xlsx");}
function exportFinance(){download(students.map(s=>({"رقم الطالب":s.number,"اسم الطالب":s.name,"القسط":s.fee,"المدفوع":s.paid,"المتبقي":Math.max(0,s.fee-s.paid),"نسبة التحصيل":s.fee?Math.round(s.paid/s.fee*100)+"%":"0%",الحالة:s.fee<=s.paid?"مكتمل":s.paid>0?"جزئي":"متأخر"})),"finance-report.xlsx");}
function exportPayments(){download(payments.map(p=>{const s=students.find(x=>x.id===p.studentId);return{"رقم الطالب":s?.number||"","اسم الطالب":s?.name||"","المبلغ":p.amount,"التاريخ":new Date(p.date).toLocaleDateString("en-GB"),"طريقة الدفع":p.method,"الملاحظة":p.note||""};}),"payments-report.xlsx");}
function exportGrades(){download(grades.map(g=>{const s=students.find(x=>x.id===g.studentId),m=subjects.find(x=>x.id===g.subjectId),total=Math.min(100,num(g.course)+num(g.final));return{"رقم الطالب":s?.number||"","اسم الطالب":s?.name||"","القسم":s?.dept||"","المرحلة":s?.stage||"","المادة":m?.name||"","الفصل":g.term||"الأول","السعي":g.course,"النهائي":g.final,"الدرجة النهائية":total,"النتيجة":total>=50?"ناجح":"راسب"};}),"grades-report.xlsx");}
function exportMaster(){const map={};grades.forEach(g=>{const s=students.find(x=>x.id===g.studentId),m=subjects.find(x=>x.id===g.subjectId);if(!s||!m)return;const key=s.id+"_"+(g.term||"الأول");map[key]??={الرقم:s.number,"اسم الطالب":s.name,القسم:s.dept,المرحلة:s.stage,"الفصل":g.term||"الأول"};map[key][m.name]=Math.min(100,num(g.course)+num(g.final));});download(Object.values(map),"master-sheet.xlsx");}
qs("#downloadTemplateBtn").onclick=()=>{
download([{الرقم:"10004","اسم الطالب":"مثال طالب","القسم":departments[0]?.name||"","المرحلة":"الأولى","الدراسة":"صباحية","القسط الكلي":1800000,"المدفوع":0}],"student-import-template.xlsx");
};
qs("#printTranscriptBtn").onclick=()=>{qs("#tStudent").innerHTML=students.map(s=>`<option value="${s.id}">${s.number} — ${s.name}</option>`).join("");qs("#tTerm").value="";qs("#transcriptDialog").showModal();};
qs("#transcriptForm").onsubmit=e=>{e.preventDefault();const studentId=qs("#tStudent").value,term=qs("#tTerm").value,s=students.find(x=>x.id===studentId);const rows=grades.filter(g=>g.studentId===studentId&&(!term||(g.term||"الأول")===term));const w=window.open("","_blank","width=900,height=750");if(!w)return;const body=rows.map(g=>{const m=subjects.find(x=>x.id===g.subjectId),total=Math.min(100,num(g.course)+num(g.final));return `<tr><td>${m?.name||""}</td><td>${g.term||"الأول"}</td><td>${g.course}</td><td>${g.final}</td><td>${total}</td><td>${total>=50?"ناجح":"راسب"}</td></tr>`}).join("");w.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><title>كشف درجات</title><style>body{font-family:Tahoma,Arial;padding:28px;color:#17243a}h2{margin-bottom:3px}.head{display:flex;justify-content:space-between;border-bottom:2px solid #17243a;padding-bottom:14px}.box{border:1px solid #ddd;padding:12px;border-radius:8px;margin:15px 0}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ccc;padding:9px;text-align:right}th{background:#f4f6f9}.pass{color:#27845c}.fail{color:#c73a3a}</style></head><body><div class="head"><div><h2>UniManager IQ</h2><div>كشف درجات طالب</div></div><div>${new Date().toLocaleDateString("en-GB")}</div></div><div class="box"><b>الطالب:</b> ${s?.name||""}<br><b>الرقم:</b> ${s?.number||""}<br><b>القسم:</b> ${s?.dept||""}<br><b>المرحلة:</b> ${s?.stage||""}</div><table><thead><tr><th>المادة</th><th>الفصل</th><th>السعي</th><th>النهائي</th><th>النهائية</th><th>النتيجة</th></tr></thead><tbody>${body||'<tr><td colspan="6">لا توجد درجات</td></tr>'}</tbody></table><script>window.print()<\/script></body></html>`);w.document.close();qs("#transcriptDialog").close();};
qs("#exportBtn").onclick=exportStudents;qs("#reportStudents").onclick=exportStudents;qs("#reportFinance").onclick=exportFinance;qs("#reportPayments").onclick=exportPayments;qs("#reportGrades").onclick=exportGrades;qs("#exportMasterBtn").onclick=exportMaster;
qs("#importBtn").onclick=()=>qs("#excelInput").click();
qs("#excelInput").onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const wb=XLSX.read(await f.arrayBuffer());const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});const mapped=rows.map((r,i)=>({id:crypto.randomUUID(),number:normalizeDigits(r["الرقم"]||r["رقم الطالب"]||r.number||i+1),name:r["اسم الطالب"]||r["الاسم"]||r.name||"",dept:r["القسم"]||r.dept||departments[0]?.name||"",stage:r["المرحلة"]||r.stage||"الأولى",study:r["الدراسة"]||r.study||"صباحية",fee:num(r["القسط الكلي"]||r["القسط"]||r.fee),paid:num(r["المدفوع"]||r.paid)})).filter(x=>x.name);if(!mapped.length)throw new Error("لم يتم العثور على صفوف صالحة");students=[...students,...mapped];save();alert(`تم استيراد ${mapped.length} طالب بنجاح`);}catch(err){alert("تعذر استيراد الملف: "+err.message)}e.target.value="";};
renderAll();