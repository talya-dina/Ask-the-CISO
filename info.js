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

// עדכון השאלות והנושאים לפי המסמך המצורף
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
    // עדכון נושא המייל עם תגית ייחודית לאוטומציה
    const fullSubject = `OFIRSEC Security (ID: ${uniqueId}) - ${type.subject} [SEC-REQ]`;
    const tableHtml = generateCyberTable(type);

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["info@ofirsec.co.il"],
        subject: fullSubject,
        htmlBody: tableHtml
    });
}

function generateCyberTable(type) {
    const rows = type.questions.map(q => `
        <tr>
            <td style="border: 1px solid #e0e0e0; padding: 12px; background-color: #fcfcfc; color: #333333; font-weight: bold; width: 40%; text-align: right; font-size: 14px;">${q}:</td>
            <td style="border: 1px solid #e0e0e0; padding: 12px; background-color: #ffffff; text-align: right;"></td>
        </tr>
    `).join("");

    return `
        <div dir="rtl" style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; max-width: 650px; color: #333; line-height: 1.6; margin: 0 auto; text-align: right;">
            <div style="background-color: #f8f9fa; padding: 20px; border: 1px solid #e0e0e0; border-bottom: 5px solid #0078d4; border-radius: 8px 8px 0 0; text-align: right;">
                <h2 style="margin: 0; font-size: 18px; color: #0078d4;">טופס בקשה: ${type.label.replace(/[^\u0590-\u05FF\s]/g, '').trim()}</h2>
                <p style="margin: 5px 0 0 0; font-size: 13px; color: #666;">אנא מלא את הפרטים בטבלה מטה והשב למייל זה.</p>
                <p style="margin: 10px 0 0 0; font-size: 14px; color: #d93025; font-weight: bold; border-top: 1px dashed #d93025; padding-top: 10px;">
                    חובה למלא את כלל הסעיפים בטבלה על מנת שיתקבל מענה לבקשה.
                </p>
            </div>
            
            <table dir="rtl" style="width: 100%; border-collapse: collapse; border: 1px solid #e0e0e0;">
                ${rows}
            </table>

            <div style="margin-top: 30px; text-align: center; border-top: 1px solid #eeeeee; padding-top: 20px;">
                <p style="margin: 0; font-size: 16px; font-weight: bold; color: #001529;">תודה רבה על שיתוף הפעולה!</p>
                <p style="margin: 5px 0 15px 0; color: #0078d4; font-size: 14px;">צוות אבטחת מידע OFIRSEC</p>
                <img src="https://ofirsec.co.il/wp-content/uploads/2024/06/logo-big-cyber-1-1-768x336.png" alt="OFIRSEC Logo" style="width: 250px; height: auto; display: block; margin: 0 auto;">
            </div>
        </div>
    `;
}
