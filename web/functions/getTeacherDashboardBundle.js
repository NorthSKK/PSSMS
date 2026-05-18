const getSystemConfig = require('./getSystemConfig');
const getCalendarEvents = require('./getCalendarEvents');
const { getTeacherTimetableWithStatus } = require('./timetable');
const { getTeacherAtRiskDashboard } = require('./attendanceReport');

function section(fn) {
  return fn().then(data => ({ ok: true, data })).catch(e => ({ ok: false, error: e.message }));
}

module.exports = async function getTeacherDashboardBundle([teacherId, term, year]) {
  const config = await getSystemConfig();
  const t = term || config.term;
  const y = year || config.year;

  const [timetable, calendarEvents, atRiskDashboard] = await Promise.all([
    section(() => getTeacherTimetableWithStatus([teacherId])),
    section(() => getCalendarEvents()),
    section(() => getTeacherAtRiskDashboard([teacherId, t, y])),
  ]);

  return { ts: Date.now(), timetable, calendarEvents, atRiskDashboard };
};
