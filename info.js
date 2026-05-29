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
        id: "open_source", 
        label: "בקשה לקוד פתוח", 
        subject: "בקשה לאישור ושימוש בתוכנת קוד פתוח (Open Source)", 
        questions: [
            "שם המערכת", 
            "סוג המערכת", 
            "כתובת האתר הרשמי", 
            "דף ההורדה", 
            "מטרת המערכת", 
            "פירוט מורחב על המערכת", 
            "פירוט הבקשה והשימוש המתבקש במערכת", 
            "היכן תותקן המערכת", 
            "הרשאות גישה למערכת", 
            "סוג רישוי", 
            "מספר עדכונים", 
            "מספר גרסה", 
            "תאריך עדכון אחרון", 
            "קריאות פתוחות", 
            "קריאות סגורות"
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

// פונקציית עיצוב המייל המלאה והמעודכנת - פונטים יציבים ואפקט "וואו" מובטח בכל הגרסאות
function generateCyberTable(type) {
    // הזרקת פונט 'Segoe UI' וצבעי סייבר ישירות לכל תא ותא למניעת שבירה באאוטלוק דסקטופ
    const rows = type.questions.map(q => `
        <tr>
            <td style="border: 1px solid #dbe6f2 !important; padding: 14px 12px; background-color: #f4f8fc !important; color: #1f385c !important; font-weight: bold; width: 38%; min-width: 120px; text-align: right; font-size: 13.5px; vertical-align: middle; word-wrap: break-word; font-family: 'Segoe UI', Tahoma, Arial, sans-serif;">${q}</td>
            <td style="border: 1px solid #dbe6f2 !important; padding: 14px 12px; background-color: #ffffff !important; text-align: right; color: #111111 !important; font-size: 15px; width: 62%; font-family: 'Segoe UI', Tahoma, Arial, sans-serif;"></td>
        </tr>
    `).join("");

    return `
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta name="color-scheme" content="light">
        <meta name="supported-color-schemes" content="light">

        <div style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; background-color:#f4f7f9; direction: rtl; text-align: right;">
        <table cellpadding="0" cellspacing="0" border="0" style="width:100%; background-color:#f4f7f9; direction: rtl; text-align: right;">
        <tbody><tr>
        <td style="padding:40px 12px;">

        <center>

            <table cellpadding="0" cellspacing="0" border="0" style="background-color: #ffffff; width: 600px; max-width: 600px; border: 1px solid #dce6f1; border-radius: 16px; overflow: hidden; box-shadow: 0 15px 35px rgba(10, 103, 181, 0.06); margin: 0 auto; direction: rtl; text-align: right;">
                <tbody>
                    <tr>
                        <td height="6" style="background-color: #0a67b5; font-size: 1px; line-height: 1px;">&nbsp;</td>
                    </tr>
                    <tr>
                    <td style="padding: 30px 25px; direction: rtl; text-align: right;">
                        
                        <table cellpadding="0" cellspacing="0" border="0" style="width:100%; direction: rtl; text-align: right;">
                            <tr>
                                <td style="padding-bottom: 4px;">
                                    <h2 style="margin: 0; font-size: 21px; color: #05335c !important; font-weight: 700; line-height: 1.3; font-family: 'Segoe UI', Tahoma, Arial, sans-serif;">
                                        טופס בקשה: ${type.label.replace(/[^\u0590-\u05FF\s]/g, '').trim()}
                                    </h2>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding-bottom: 12px;">
                                    <table cellpadding="0" cellspacing="0" border="0" style="width: 45px; height: 3px; background-color: #5bb8c9; border-radius: 2px;"><tr><td></td></tr></table>
                                </td>
                            </tr>
                        </table>

                        <p style="margin: 0 0 16px 0; font-size: 14px; color: #505050 !important; line-height: 1.5; font-family: 'Segoe UI', Tahoma, Arial, sans-serif;">
                            אנא מלאו את הפרטים בשדות מטה והשיבו למייל זה.
                        </p>
                        
                        <p style="margin: 0 0 24px 0; font-size: 13px; color: #d93025 !important; font-weight: bold; border-top: 1px dashed rgba(217, 48, 37, 0.15); padding-top: 12px; line-height: 1.4; font-family: 'Segoe UI', Tahoma, Arial, sans-serif;">
                            חובה למלא את כלל הסעיפים על מנת שיתקבל מענה לבקשה.
                        </p>
                        
                        <div style="border: 1px solid #cfe0f2 !important; border-radius: 10px; overflow: hidden; width: 100%;">
                            <table dir="rtl" style="width: 100%; border-collapse: collapse; table-layout: fixed; background-color: #ffffff !important;">
                                <tbody>
                                    ${rows}
                                </tbody>
                            </table>
                        </div>
                    </td>
                </tr>
            </tbody></table>

            <table cellpadding="0" cellspacing="0" border="0" style="width: 600px; max-width: 600px; margin-top: 25px; margin-left: auto; margin-right: auto; direction: rtl;">
                <tbody><tr>
                    <td style="text-align: center;">
                        <p style="margin: 0; font-size: 15px; font-weight: bold; color: #001529 !important; font-family: 'Segoe UI', Tahoma, Arial, sans-serif;">תודה רבה על שיתוף הפעולה!</p>
                        <p style="margin: 4px 0 12px 0; color: #0a67b5 !important; font-size: 13.5px; font-weight: bold; font-family: 'Segoe UI', Tahoma, Arial, sans-serif;">צוות אבטחת מידע OFIRSEC</p>
                        <img src="https://ofirsec.co.il/wp-content/uploads/2024/06/logo-big-cyber-1-1-1024x448.png" alt="OFIRSEC Logo" width="150" style="width: 150px; height: auto; display: block; margin: 0 auto; border: 0;">
                    </td>
                </tr>
            </tbody></table>

        </center>

        </td>
        </tr>
        </tbody></table>
        </div>
    `;
}
