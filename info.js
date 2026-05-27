function initializeApp() {
    renderButtons();
}

Office.onReady((info) => {
    if (info.host) {
        initializeApp();
    }
});

if (!window.officeInitialized && (window.location.host.includes('github.io') || window.location.host.includes('localhost'))) {
    if (document.readyState === "complete" || document.readyState === "interactive") {
        initializeApp();
    } else {
        document.addEventListener("DOMContentLoaded", initializeApp);
    }
}

// רשימת השאלות והנושאים
const requestTypes = [
    { 
        id: "internet", 
        label: "יציאה מיוחדת לאינטרנט", 
        subject: "בקשה ליציאה מיוחדת של רכיב לאינטרנט", 
        questions: [
            "שם הרכיב שנדרש לצאת לאינטרנט", 
            "כתובות הIP של הרכיב הנדרש לצאת (במידה וקיימת כתובת קבועה)", 
            "רשימת האתרים/ כתובות ה IP אליהם השירות נדרש לצאת", 
            "תיאור הצורך ביציאה לאינטרנט (במידה ונדרש לכלל האינטרנט יש לנמק)", 
            "פורט נדרש ליציאה לעולם", 
            "הסבר על סיבת הפורט הספציפי"
        ] 
    },
    { 
        id: "privileges", 
        label: "הרשאות פריבילגיות", 
        subject: "בקשה למתן הרשאות פריבילגיות ברשת", 
        questions: [
            "מטרת ההרשאה והגדרות התפקיד", 
            "לאיזה קבוצות חזקות המשתמש יצטרף", 
            "מייל המשתמש", 
            "תפקיד", 
            "באיזה סביבה המשתמש ימצא (פנימית / עננית / אפליקטיבית)"
        ] 
    },
    { 
        id: "generic", 
        label: "משתמש גנרי / סרביס", 
        subject: "בקשה לפתיחת משתמש גנרי (לרבות סרביס)", 
        questions: [
            "שם המשתמש הגנרי", 
            "תפקיד המשתמש / סיבה לפתיחתו", 
            "הרשאות נדרשות", 
            "עמדות או מערכות אליהם יהיה רשאי לגשת"
        ] 
    },
    { 
        id: "software", 
        label: "תוכנה / מערכת חדשה", 
        subject: "בקשה לתוכנה/תוסף /מערכת חדשה", 
        questions: [
            "שם התוכנה", 
            "שם החברה", 
            "קישור לאתר הרלוונטי", 
            "סוג התוכנה (מקומית / עננית / לא ידוע)",
            "האם נדרש רשיון/חינמי",
            "מטרת התוכנה",
            "אילו משתמשים צפויים להשתמש בתוכנה",
            "תיאור הצורך בתוכנה",
        ] 
    },
    { 
        id: "gritta", 
        label: "תיעוד גריטה", 
        subject: "תיעוד גריטה", 
        questions: [
            "תאריך ביצוע הגריטה", 
            "שם מבצע הגריטה", 
            "מקום ביצוע הגריטה", 
            "הרכיב שנגרט", 
            "המידע שהיה על הרכיב"
        ] 
    },
    { 
        id: "survey", 
        label: "פרסום טופס או סקר", 
        subject: "פרסום טופס או סקר", 
        questions: [
            "שם המחלקה", 
            "תאריך פתיחת הטופס", 
            "תאריך סגירת הטופס", 
            "שם הטופס / סקר", 
            "הנתונים שיאספו בטופס / סקר", 
            "כתובות לטופס / סקר", 
            "מערכת עליה מתבסס הטופס / סקר"
        ] 
    },
    { 
        id: "general", 
        label: "אישור כללי אחר", 
        subject: "אישור כללי/אחר", 
        questions: [
            "פירוט פרטי האישור והצורך בו"
        ] 
    }
];

function renderButtons() {
    const list = document.getElementById("button-list");
    if (!list) return;
    list.innerHTML = "";
    requestTypes.forEach(type => {
        const btn = document.createElement("button");
        btn.className = "request-btn";
        btn.innerHTML = `<span>${type.label}</span>`;
        btn.onclick = () => openNewEmail(type);
        list.appendChild(btn);
    });
}

function openNewEmail(type) {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) {
        alert("הפעולה זמינה רק מתוך Outlook.");
        return;
    }

    const uniqueId = Date.now();
    const fullSubject = `OFIRSEC Security (ID: ${uniqueId}) - ${type.subject} [SEC-REQ]`;
    const tableHtml = generateCyberTable(type);

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["info@ofirsec.co.il"],
        subject: fullSubject,
        htmlBody: tableHtml
    });
}

// פונקציית עיצוב המייל שמתאימה בול לכרטיס הלבן והתכלת האחרון
function generateCyberTable(type) {
    // עיצוב מחדש של השורות - ללא קוביות אפורות, מראה נקי כמו רשימת פרטים
    const rows = type.questions.map(q => `
        <tr style="border-bottom: 1px solid #e1e6eb;">
            <td style="padding: 16px 10px; color: #605e5c; font-weight: 500; width: 40%; text-align: right; font-size: 15px; vertical-align: top;">${q}:</td>
            <td style="padding: 16px 10px; text-align: right; color: #0a67b5; font-weight: 600; font-size: 16px;">
                [הקלידו כאן את המענה]
            </td>
        </tr>
    `).join("");

    return `
        <div dir="rtl" style="background-color: #f0f4f8; padding: 30px 20px; font-family: 'Segoe UI', Tahoma, Arial, sans-serif; max-width: 650px; margin: 0 auto; direction: rtl; text-align: right; border-radius: 20px;">
            
            <div style="background-color: #ffffff; border-radius: 16px; border: 1px solid rgba(255, 255, 255, 0.9); box-shadow: 0 15px 35px rgba(10, 103, 181, 0.08), 0 3px 10px rgba(0, 0, 0, 0.02); overflow: hidden; padding: 30px 25px;">
                
                <h2 style="margin: 0 0 8px 0; font-size: 20px; color: #0a67b5; font-weight: 700;">
                    טופס בקשה: ${type.label.replace(/[^\u0590-\u05FF\s]/g, '').trim()}
                </h2>
                <p style="margin: 0; font-size: 14px; color: #605e5c; line-height: 1.5;">
                    אנא מלאו את הפרטים בשדות מטה והשיבו למייל זה.
                </p>
                
                <p style="margin: 15px 0 25px 0; font-size: 13.5px; color: #d93025; font-weight: bold; border-top: 1px dashed rgba(217, 48, 37, 0.2); padding-top: 12px;">
                    ⚠️ חובה למלא את כלל הסעיפים על מנת שיתקבל מענה לבקשה.
                </p>
                
                <table dir="rtl" style="width: 100%; border-collapse: collapse; margin-bottom: 10px;">
                    <tbody>
                        ${rows}
                    </tbody>
                </table>
            </div>

            <div style="margin-top: 30px; text-align: center;">
                <p style="margin: 0; font-size: 16px; font-weight: bold; color: #001529;">תודה רבה על שיתוף הפעולה!</p>
                <p style="margin: 5px 0 15px 0; color: #0a67b5; font-size: 14px; font-weight: bold;">צוות אבטחת מידע OFIRSEC</p>
                <img src="https://ofirsec.co.il/wp-content/uploads/2024/06/logo-big-cyber-1-1-768x336.png" alt="OFIRSEC Logo" style="width: 170px; height: auto; display: block; margin: 0 auto; filter: drop-shadow(0 4px 8px rgba(10, 103, 181, 0.05)); opacity: 0.95;">
            </div>
        </div>
    `;
}
