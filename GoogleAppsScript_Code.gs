/** 환경 성찰 검사 v4 / Google Sheets bound Apps Script. */
const DATA_SHEET='설문응답';
const PAIRS_SHEET='전후비교';
const DASH_SHEET='종합통계';
const HEADERS=['접수시각','이름/별명','학생식별키','검사구분','연령구분','환경총점','일관성','재확인필요','레벨','성경인물','재활용 이해','환경 성찰','학교/집 일관성','실제 분리배출','자발적 확산',...Array.from({length:20},(_,i)=>'Q'+(i+1)),'클라이언트시각','버전','반','학년'];
function onOpen(){SpreadsheetApp.getUi().createMenu('환경 검사 관리').addItem('전후 비교·통계 갱신','buildReport').addToUi();}
function safeCell(v,max=80){let s=String(v??'').slice(0,max);return /^[=+\-@\t\r]/.test(s)?"'"+s:s;}
function doPost(e){
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const x=JSON.parse((e.postData&&e.postData.contents)||'{}');
  if(!Array.isArray(x.answers)||x.answers.length!==20||x.answers.some(v=>!Number.isInteger(v)||v<1||v>5))throw Error('Invalid answers');
  const phase=String(x.phase||'');if(!['사전','사후'].includes(phase))throw Error('Invalid phase');
  const name=String(x.name||'').trim().normalize('NFC'),className=String(x.className||'').trim(),grade=String(x.grade||'').trim();if(!name||name.length>24||!['로뎀나무','감람나무','에셀나무','종려나무'].includes(className)||!Array.from({length:12},(_,i)=>String(i+1)).includes(grade))throw Error('Invalid identity');const code=[className,grade,name.replace(/\s+/g,'')].join('|');
  const ss=SpreadsheetApp.getActiveSpreadsheet();const sh=ss.getSheetByName(DATA_SHEET)||ss.insertSheet(DATA_SHEET);
  if(sh.getLastRow()===0)sh.appendRow(HEADERS);
  const rawScores=Array.isArray(x.scores)?x.scores:[];if(rawScores.length!==5)throw Error('Invalid scores');
  const scores=rawScores.map(Number);const num=Number(x.score),consistency=Number(x.consistency);
  if([num,consistency,...scores].some(z=>!Number.isFinite(z)||z<0||z>100))throw Error('Invalid scores');
  sh.appendRow([new Date(),safeCell(x.name,24),safeCell(code,16),phase,safeCell(x.ageGroup,10),num,consistency,consistency<=70,consistency<=70?'':Number(x.level),consistency<=70?'야곱':safeCell(x.character,36),...scores,...x.answers,safeCell(x.submittedAt,40),safeCell(x.version,32),className,Number(grade)]);
  return ContentService.createTextOutput('OK');
 }catch(err){return ContentService.createTextOutput('ERROR');}
 finally{lock.releaseLock();}
}
function buildReport(){
 const ss=SpreadsheetApp.getActiveSpreadsheet(),src=ss.getSheetByName(DATA_SHEET);
 if(!src||src.getLastRow()<2){SpreadsheetApp.getUi().alert('먼저 사전·사후 응답을 저장해 주세요.');return;}
 const data=src.getDataRange().getValues().slice(1), perCode={};
 data.forEach((row,idx)=>{
   const code=String(row[2]).trim(),phase=String(row[3]),date=row[0];
   if(!code.includes('|')||!['사전','사후'].includes(phase))return;
   if(!perCode[code])perCode[code]={};
   const old=perCode[code][phase];
   // 같은 코드/시점의 중복 응답: 가장 최근 기록 사용
   if(!old||new Date(date).getTime()>=new Date(old[0]).getTime())perCode[code][phase]=row;
 });
 const valid=[],flags=[],unmatched=[];let duplicateCodes=0;
 Object.entries(perCode).forEach(([code,entries])=>{
   const pre=entries['사전'],post=entries['사후'];
   if(!pre||!post){unmatched.push(code);return;}
   const pair=[code,safeCell(post[1]||pre[1],24),String(pre[4]),Number(pre[5]),Number(post[5]),Number(post[5])-Number(pre[5]),pre[8],post[8],Number(post[8])-Number(pre[8]),
     ...(Array.from({length:5},(_,i)=>[Number(pre[10+i]),Number(post[10+i]),Number(post[10+i])-Number(pre[10+i])]).flat()),
     Number(pre[6]),Number(post[6]),pre[7]===true||post[7]===true?'재확인':'정상',new Date(pre[0]),new Date(post[0])];
   if(pair[26]==='재확인')flags.push(pair);else valid.push(pair);
 });
 const headers=['학생식별키','이름/별명','연령구분','사전총점','사후총점','점수변화','사전레벨','사후레벨','레벨변화',...['재활용 이해','환경 성찰','학교/집 일관성','실제 분리배출','자발적 확산'].flatMap(n=>[n+' 사전',n+' 사후',n+' 변화']),'사전일관성','사후일관성','분석상태','사전응답시각','사후응답시각'];
 const sh=ss.getSheetByName(PAIRS_SHEET)||ss.insertSheet(PAIRS_SHEET);sh.clear();sh.getRange(1,1,1,headers.length).setValues([headers]);
 const all=[...valid,...flags];if(all.length)sh.getRange(2,1,all.length,headers.length).setValues(all);
 sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#dbeee4');sh.autoResizeColumns(1,Math.min(headers.length,9));
 const report=ss.getSheetByName(DASH_SHEET)||ss.insertSheet(DASH_SHEET);report.clear();
 const n=valid.length;const mean=(a)=>a.length?a.reduce((s,x)=>s+x,0)/a.length:null;
 const pre=valid.map(r=>r[3]),post=valid.map(r=>r[4]),deltas=valid.map(r=>r[5]);
 const improve=valid.filter(r=>r[5]>0).length,decline=valid.filter(r=>r[5]<0).length,stable=valid.filter(r=>r[5]===0).length;
 const summary=[['환경 성찰 검사 · 사전·사후 분석','값'],['분석 대상: 반·학년·이름으로 연결되고 두 번 모두 일관성 70 초과인 응답',n],['매칭되지만 일관성 재확인 필요한 학생',flags.length],['사전 또는 사후만 있어 미매칭된 학생',unmatched.length],['사전 평균점수',mean(pre)],['사후 평균점수',mean(post)],['평균 변화점수',mean(deltas)],['점수 상승 학생 수',improve],['변화 없음 학생 수',stable],['점수 하락 학생 수',decline],['상승 학생 비율',n?improve/n:null],['하락 학생 비율',n?decline/n:null],['등급 상승 학생 수',valid.filter(r=>r[8]>0).length],['등급 하락 학생 수',valid.filter(r=>r[8]<0).length],['등급 유지 학생 수',valid.filter(r=>r[8]===0).length],['참고: 두 시점의 응답 일관성 기준','각 70점 초과'],['주의','점수 변화는 인과관계 입증이 아니며, 소규모 표본에서는 불확실성이 큽니다.']];
 report.getRange(1,1,summary.length,2).setValues(summary);report.getRange('A1:B1').setFontWeight('bold').setBackground('#dbeee4');report.getRange('B11:B12').setNumberFormat('0.0%');report.getRange('B5:B7').setNumberFormat('0.0');
 const domainHeads=['영역','사전 평균','사후 평균','변화'];report.getRange(19,1,1,4).setValues([domainHeads]).setFontWeight('bold').setBackground('#dbeee4');
 const domainNames=['재활용 이해','환경 성찰','학교/집 일관성','실제 분리배출','자발적 확산'];
 const domainRows=domainNames.map((name,i)=>[name,mean(valid.map(r=>r[9+i*3])),mean(valid.map(r=>r[10+i*3])),mean(valid.map(r=>r[11+i*3]))]);
 report.getRange(20,1,5,4).setValues(domainRows);
 report.getRange('A27').setValue('짝지은 평균차이의 통계 검정 (참고용)');
 report.getRange('A28').setValue('대응표본 양측 t 검정 p값');
 if(n>=3) report.getRange('B28').setFormula(`=IFERROR(TTEST('${PAIRS_SHEET}'!D2:D${n+1},'${PAIRS_SHEET}'!E2:E${n+1},2,1),"계산 불가")`);
 else report.getRange('B28').setValue('매칭된 유효 응답 3명 이상 필요');
 report.getRange('A29').setValue('해석 주의');report.getRange('B29').setValue('p값만으로 교육적 효과·원인을 단정할 수 없습니다. 문항 타당도 및 표본 구성을 검토하세요.');
 report.autoResizeColumns(1,2);
 SpreadsheetApp.getUi().alert(`통계 갱신 완료: 유효 매칭 ${n}명, 재확인 ${flags.length}명, 미매칭 ${unmatched.length}명`);
}
