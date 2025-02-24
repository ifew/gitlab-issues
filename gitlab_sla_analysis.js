// Node.js Script for GitLab Group-Level Issues and SLA Analysis with Graphs
// Requirements: axios, moment, chartjs-node-canvas, exceljs
// Install: npm install axios moment chartjs-node-canvas exceljs



// Node.js Script for GitLab Group-Level Issues and SLA Analysis (No Charts)
// Requirements: axios, moment, exceljs
// Install: npm install axios moment exceljs

const axios = require('axios');
// const moment = require('moment');
const moment = require('moment-business-days');
const ExcelJS = require('exceljs');

// GitLab Configuration
const GITLAB_URL = 'http://gitlab.com';
const GROUP_ID = '10'; //10
const ACCESS_TOKEN = 'xxxTOKENxx';
const headers = { 'PRIVATE-TOKEN': ACCESS_TOKEN };
const PER_PAGE = 100; // Max issues per page

async function fetchAllGroupIssues() {
  let allIssues = [];
  let page = 1;
  while (true) {
    const response = await axios.get(`${GITLAB_URL}/api/v4/groups/${GROUP_ID}/issues`, {
      headers,
      params: { per_page: PER_PAGE, page }
    });
    if (response.data.length === 0) break;
    allIssues.push(...response.data);
    page++;
  }
  return allIssues;
}

function calculateWorkingHours(startTime, durationMinutes) {
  if (!moment.isMoment(startTime)) {
    startTime = moment(startTime);
  }

  let endTime = startTime.clone().add(durationMinutes, 'minutes');
  const workStartHour = 9;
  const workEndHour = 18;
  let totalWorkMinutes = 0;

  let current = startTime.clone();
  while (current.isBefore(endTime)) {
    if (current.isBusinessDay()) { // ตรวจสอบว่าเป็นวันทำการหรือไม่
      let startOfWorkDay = current.clone().hour(workStartHour).minute(0).second(0);
      let endOfWorkDay = current.clone().hour(workEndHour).minute(0).second(0);

      if (current.isBefore(startOfWorkDay)) {
        current = startOfWorkDay;
      }

      if (current.isBefore(endOfWorkDay)) {
        let minutesAvailable = Math.min(endOfWorkDay.diff(current, 'minutes'), endTime.diff(current, 'minutes'));
        totalWorkMinutes += minutesAvailable;
        current.add(minutesAvailable, 'minutes');
      }
    }

    // ไปวันทำการถัดไป
    current.add(1, 'day').startOf('day').hour(workStartHour);
  }

  return totalWorkMinutes;
}

async function fetchIssueActivities(issueId, projectId) {
  try {
    const response = await axios.get(`${GITLAB_URL}/api/v4/projects/${projectId}/issues/${issueId}/resource_label_events`, { headers });
    const statusChanges = response.data
      .filter(activity => activity.label && activity.label.name.startsWith('STATUS : ') && activity.action === 'add')
      .reduce((acc, activity) => {
        const statusName = activity.label.name.replace('STATUS : ', '');
        const activityTime = moment(activity.created_at).format('YYYY-MM-DD HH:mm:ss');
        
        if (statusName === "OPEN" && !acc[statusName]) {
          acc[statusName] = activityTime; // ใช้ตัวแรกที่พบสำหรับ "OPEN"
        } else if (statusName === "Close") {
          acc[statusName] = activityTime; // ใช้ตัวสุดท้ายที่พบสำหรับ "Close"
        } else if (statusName === "Fix") {
          acc[statusName] = activityTime; // ใช้ตัวสุดท้ายที่พบสำหรับ "Fix"
        }
        
        return acc;
      }, {});
    
    // คำนวณระยะเวลาจาก "OPEN" → "Fix", "Fix" → "Close", "OPEN" → "Close"
    let durationOpenToFix = null;
    let durationFixToClose = null;
    let durationOpenToClose = null;
    let durationOpenToFixBiz = null;
    let durationFixToCloseBiz = null;
    let durationOpenToCloseBiz = null;
    let statusLabel = "Open";
    let slaStatus = "Pass";
    let slaExceededHours = 0;
    let slaExceededMinutes = 0;

    if (statusChanges["OPEN"] && statusChanges["Fix"]) {
      const openTime = moment(statusChanges["OPEN"], "YYYY-MM-DD HH:mm:ss");
      const fixTime = moment(statusChanges["Fix"], "YYYY-MM-DD HH:mm:ss");
      durationOpenToFix = fixTime.diff(openTime, 'minutes');
    }

    if (statusChanges["Fix"] && statusChanges["Close"]) {
      const fixTime = moment(statusChanges["Fix"], "YYYY-MM-DD HH:mm:ss");
      const closeTime = moment(statusChanges["Close"], "YYYY-MM-DD HH:mm:ss");
      durationFixToClose = closeTime.diff(fixTime, 'minutes');

      if (durationFixToClose < 0) {
        durationFixToClose = null;
        durationOpenToClose = null;
        statusLabel = "Fix";
      } else {
        statusLabel = "Close";
      }
    }

    if (statusChanges["OPEN"] && statusChanges["Close"]) {
      const openTime = moment(statusChanges["OPEN"], "YYYY-MM-DD HH:mm:ss");
      const closeTime = moment(statusChanges["Close"], "YYYY-MM-DD HH:mm:ss");
      durationOpenToClose = closeTime.diff(openTime, 'minutes');
    }

    if (statusChanges["Fix"] && !statusChanges["Close"]) {
      statusLabel = "Fix";
    }
    
    if (statusChanges['OPEN'] && durationOpenToFix != null) {
      durationOpenToFixBiz = calculateWorkingHours(
        moment(statusChanges['OPEN'], 'YYYY-MM-DD HH:mm:ss'),
        durationOpenToFix
      );
    }

    if (statusChanges['Fix'] && durationFixToClose != null) {
      durationFixToCloseBiz = calculateWorkingHours(
        moment(statusChanges['Fix'], 'YYYY-MM-DD HH:mm:ss'),
        durationFixToClose
      );
    }

    if (statusChanges['OPEN'] && durationOpenToClose != null) {
      durationOpenToCloseBiz = calculateWorkingHours(
        moment(statusChanges['OPEN'], 'YYYY-MM-DD HH:mm:ss'),
        durationOpenToClose
      );
    }
    
    // ดึงค่าของ Develop Label
    const developActivity = response.data.find(activity => activity.label && activity.label.name.startsWith('Develop : '));
    if (developActivity) {
      developLabel = developActivity.label.name.replace('Develop : ', '');
    }

    const priorityActivity = response.data.find(activity => activity.label && activity.label.name.startsWith('Priority : '));
    if (priorityActivity) {
      priorityLabel = priorityActivity.label.name.replace('Priority : ', '');
    }

    const prioritySLA = { "Urgent": 4, "High": 8, "Medium": 24, "Low": 40 };

    if (developLabel === 'Bug' && durationOpenToFixBiz !== null && priorityLabel in prioritySLA) {
      const slaLimit = prioritySLA[priorityLabel] * 60; // SLA limit เป็นนาที
      if (durationOpenToFixBiz > slaLimit) {
        slaStatus = 'Fail';
        const exceeded = durationOpenToFixBiz - slaLimit;
        slaExceededHours = Math.floor(exceeded / 60);
        slaExceededMinutes = exceeded % 60;
      }
    }
    
    return { 
      statusChanges: JSON.stringify(statusChanges), 
      durationOpenToFix, 
      durationFixToClose, 
      durationOpenToClose,
      statusLabel,
      developLabel,
      slaStatus,
      slaExceededHours,
      slaExceededMinutes,
      durationOpenToFixBiz,
      durationFixToCloseBiz,
      durationOpenToCloseBiz
    };
  } catch (error) {
    if (error.response && error.response.status === 404) {
      console.warn(`Activities not found for issue ${issueId} in project ${projectId}`);
      return { statusChanges: '{}', durationOpenToFix: null, statusLabel: "Open", developLabel: null, priorityLabel: null, slaStatus: "Pass", slaExceededHours: 0, slaExceededMinutes: 0, durationOpenToFixBiz: null, durationFixToCloseBiz: null, durationOpenToCloseBiz: null };
    }
    throw error;
  }
}

// ─────────────────────────────────────────────────────────
//                      EXPORT TO EXCEL
// ─────────────────────────────────────────────────────────
async function exportResultsToExcel(issues) {
  const workbook = new ExcelJS.Workbook();
  // ───── SHEET 1: DETAIL REPORT ──────────
  const worksheet = workbook.addWorksheet('Issues Report');
  worksheet.columns = [
    { header: 'Issue ID', key: 'id' },
    { header: 'Issue IID', key: 'iid' },
    { header: 'Project ID', key: 'project_id' },
    { header: 'Title', key: 'title' },
    { header: 'State', key: 'state' },
    { header: 'Author', key: 'author' },
    { header: 'Assignee', key: 'assignee' },
    { header: 'Priority', key: 'priority' },
    { header: 'Status Label', key: 'status_label' },
    { header: 'Develop Label', key: 'develop_label' },
    { header: 'Issue Year', key: 'issue_year' },
    { header: 'Issue Month', key: 'issue_month' },
    { header: 'Issue Quarter', key: 'issue_quarter' },
    { header: 'Created At', key: 'created_at' },
    { header: 'Updated At', key: 'updated_at' },
    { header: 'Labels', key: 'labels' },
    { header: 'Status Changes (JSON)', key: 'status_changes' },
    { header: 'Duration (Open → Fix) (minutes)', key: 'duration_open_to_fix' },
    { header: 'Duration (Fix → Close) (minutes)', key: 'duration_fix_to_close' },
    { header: 'Duration (Open → Close) (minutes)', key: 'duration_open_to_close' },
    {
      header: 'Duration (Open → Fix) (business min)',
      key: 'durationOpenToFixBiz'
    },
    {
      header: 'Duration (Fix → Close) (business min)',
      key: 'durationFixToCloseBiz'
    },
    {
      header: 'Duration (Open → Close) (business min)',
      key: 'durationOpenToCloseBiz'
    },
    { header: 'SLA Status', key: 'slaStatus' },
    { header: 'SLA Exceeded Hours', key: 'slaExceededHours' },
    { header: 'SLA Exceeded Minutes', key: 'slaExceededMinutes' }
  ];

  // เก็บข้อมูลทั้งหมดเป็น array เพื่อใช้สรุปรายงานต่อ
  const finalData = [];
  
  for (const issue of issues) {
    const { statusChanges, durationOpenToFix, durationFixToClose, durationOpenToClose, statusLabel, developLabel, slaStatus, slaExceededHours, slaExceededMinutes, durationOpenToFixBiz, durationFixToCloseBiz,durationOpenToCloseBiz } = await fetchIssueActivities(issue.iid, issue.project_id);
    const authorName = issue.author ? issue.author.name : 'N/A';
    const assigneeName = issue.assignees && issue.assignees.length > 0 ? issue.assignees.map(a => a.name).join(', ') : 'Unassigned';
    const priorityLabel = issue.labels.find(label => label.startsWith('Priority : '))?.replace('Priority : ', '') || 'N/A';
    const issueYear = moment(issue.created_at).year();
    const issueMonth = moment(issue.created_at).format('MMMM');
    const issueQuarter = `Q${moment(issue.created_at).quarter()}`;
    
    const rowData = {
      id: issue.id,
      iid: issue.iid,
      project_id: issue.project_id,
      title: issue.title,
      state: issue.state,
      status_label: statusLabel,
      develop_label: developLabel,
      author: authorName,
      assignee: assigneeName,
      priority: priorityLabel,
      issue_year: issueYear,
      issue_month: issueMonth,
      issue_quarter: issueQuarter,
      created_at: moment(issue.created_at).format('YYYY-MM-DD HH:mm:ss'),
      updated_at: moment(issue.updated_at).format('YYYY-MM-DD HH:mm:ss'),
      labels: issue.labels ? issue.labels.join(', ') : '',
      status_changes: statusChanges,
      duration_open_to_fix: durationOpenToFix,
      duration_fix_to_close: durationFixToClose,
      duration_open_to_close: durationOpenToClose,
      durationOpenToFixBiz: durationOpenToFixBiz,
      durationFixToCloseBiz: durationFixToCloseBiz,
      durationOpenToCloseBiz: durationOpenToCloseBiz,
      slaStatus: slaStatus,
      slaExceededHours: slaExceededHours,
      slaExceededMinutes: slaExceededMinutes
    };

    // เขียนลง Sheet 1
    worksheet.addRow(rowData);

    // เก็บไว้สรุปต่อ
    finalData.push(rowData);
  }

// ──────────────────────────────────────────
  // Sheet 2: Summary Report (Group by Year + Assignee + Priority + Develop Label)
  // ──────────────────────────────────────────
  const summarySheet = workbook.addWorksheet('Summary Report');
  summarySheet.columns = [
    { header: 'Year', key: 'year' },
    { header: 'Assignee', key: 'assignee' },
    { header: 'Priority', key: 'priority' },
    { header: 'Develop Label', key: 'develop' },
    { header: 'Total Issues', key: 'total_issues' },
    { header: 'Pass Issues', key: 'pass_count' },
    { header: 'Fail Issues', key: 'fail_count' },
    { header: 'Pass (%)', key: 'pass_percent' },
    { header: 'Fail (%)', key: 'fail_percent' },
  ];

  // aggregator
  const aggregator = {};

  for (const row of finalData) {
    const year = row.issue_year; // group by year
    const assignee = row.assignee;
    const priority = row.priority;
    const develop = row.develop_label || 'N/A';
    const slaStatus = row.slaStatus; // Pass / Fail

    const key = `${year}||${assignee}||${priority}||${develop}`;
    if (!aggregator[key]) {
      aggregator[key] = {
        year,
        assignee,
        priority,
        develop,
        total: 0,
        pass: 0,
        fail: 0,
      };
    }

    aggregator[key].total += 1;
    if (slaStatus === 'Fail') aggregator[key].fail += 1;
    else aggregator[key].pass += 1;
  }

  // สร้าง array แล้ว sort
  const summaryArray = Object.values(aggregator);

  // Sort: ปีล่าสุด -> เก่าสุด, ถ้าปีเท่ากัน เรียง assignee จากน้อย -> มาก
  // หมายเหตุ: 'ปีล่าสุด' = 'ปีเลขมาก' -> จึงต้อง sort (b.year - a.year)
  // จากนั้น sort assignee ascending
  summaryArray.sort((a, b) => {
    if (b.year !== a.year) {
      return b.year - a.year; // year descending
    }
    // if year is the same => sort by assignee ascending
    return a.assignee.localeCompare(b.assignee);
  });

  // เขียนลง Sheet Summary
  for (const info of summaryArray) {
    const passPercent = info.total ? (info.pass / info.total) * 100 : 0;
    const failPercent = info.total ? (info.fail / info.total) * 100 : 0;
    summarySheet.addRow({
      year: info.year,
      assignee: info.assignee,
      priority: info.priority,
      develop: info.develop,
      total_issues: info.total,
      pass_count: info.pass,
      fail_count: info.fail,
      pass_percent: passPercent.toFixed(2),
      fail_percent: failPercent.toFixed(2),
    });
  }

  
  await workbook.xlsx.writeFile('group_issues_report.xlsx');
  console.log('Issues report saved as group_issues_report.xlsx');
}

(async () => {
  console.log('Fetching all issues from group...');
  const allIssues = await fetchAllGroupIssues();
  console.log(`Fetched ${allIssues.length} issues.`);
  await exportResultsToExcel(allIssues);
})();