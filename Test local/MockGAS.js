// ตรวจสอบว่าไม่ได้รันอยู่บนเซิร์ฟเวอร์ของ Google
if (typeof google === 'undefined') {
  console.warn("⚠️ กำลังรันในโหมด Local Development (Mock GAS)");

  window.google = {
    script: {
      run: {
        // --- ระบบรองรับ Callback ของ GAS ---
        withFailureHandler: function(failCallback) {
          this._failCallback = failCallback;
          return this; // เพื่อให้ Chain method ต่อได้
        },
        withSuccessHandler: function(successCallback) {
          this._successCallback = successCallback;
          return this; // เพื่อให้ Chain method ต่อได้
        },

        // ==========================================
        // จำลองฟังก์ชันฝั่งเซิร์ฟเวอร์ (แก้ไขให้ตรงกับ Code.js)
        // ==========================================
        
        // 1. ระบบ Login
        checkLogin: function(user, pass) {
          console.log(`[Mock] กำลังตรวจสอบ Login: ${user}`);
          const successCb = this._successCallback;
          
          setTimeout(() => { // จำลองเน็ตหน่วง 1 วินาที
            if (user === 'admin' || user === 'teacher') {
              successCb({
                status: "success",
                id: "T001",
                name: "ครูน๊อต ศิกษก (Local)",
                role: "TEACHER",
                currentTerm: "1",
                currentYear: "2568"
              });
            } else {
              successCb({ status: "error", message: "รหัสผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง (Mock)" });
            }
          }, 1000);
        },

        // 2. โหลดรายการวิชา
        getTeacherSubjects: function(teacherId, role, term, year) {
          const successCb = this._successCallback;
          setTimeout(() => {
            successCb([
              ["ว31101", "วิทยาศาสตร์กายภาพ", "4/1", "รายวิชาพื้นฐาน", "3"],
              ["ว31101", "วิทยาศาสตร์กายภาพ", "4/2", "รายวิชาพื้นฐาน", "3"]
            ]);
          }, 500);
        },

        // 3. โหลดโครงสร้างคะแนน ปพ.5 (All-in-One)
        getAllInOneScoreGridData: function(subjectCode, className, term, year) {
          const successCb = this._successCallback;
          setTimeout(() => {
            // โยนข้อมูลจำลอง (Mock Data) กลับไปให้หน้าเว็บ
            successCb({
              config: { 
                ratio: "70:10:20", 
                indicators: [
                  { name: "ใบงานที่ 1", score: 10 },
                  { name: "ทดสอบย่อย", score: 20 },
                  { name: "ชิ้นงาน", score: 40 }
                ] 
              },
              students: [
                ["12345", "1", "นายสมชาย เรียนดี"],
                ["12346", "2", "นางสาวสมหญิง ตั้งใจ"]
              ],
              existingScores: {
                "12345_ind_ใบงานที่ 1": 9,
                "12345_midterm": 8
              },
              existingQuals: {},
              attStats: {
                "12345": 95, // % เวลาเรียน
                "12346": 78  // ติด มส. (ต่ำกว่า 80%)
              }
            });
          }, 800);
        },

        // 4. บันทึกคะแนน
        saveAllInOneWithConfig: function(payload) {
          console.log("[Mock] ได้รับคำสั่งบันทึกคะแนน:", payload);
          const successCb = this._successCallback;
          setTimeout(() => {
            successCb({ status: "success", message: "บันทึกข้อมูลจำลองสำเร็จ!" });
          }, 1200);
        }

      }
    }
  };
}