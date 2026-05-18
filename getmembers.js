
const database = require('../../database');
const XLSX = require('xlsx'); // استدعاء مكتبة الاكسل التي قمنا بتثبيتها

module.exports = {
    name: 'getmembers', // تم تغيير الاسم ليكون معبراً عن الوظيفة
    aliases: ['الأعضاء', 'سحب_الأرقام', 'members'], // اختصارات للأمر بالعربي والانجليزي
    category: 'admin',
    description: 'نسخ أرقام جميع أعضاء الجروب وتحويلها إلى ملف Excel (.xlsx)',
    usage: '.getmembers',
    groupOnly: true,    // يجب أن يكون true لأننا نسحب أعضاء جروب
    adminOnly: true,    // يفضل تركه true حتى لا يقوم أي عضو بسحب الأرقام لحماية الخصوصية
    botAdminNeeded: false,

    async execute(sock, msg, args, extra) {
        try {
            const groupJid = msg.key.remoteJid;

            // 1. إرسال رسالة تفيد ببدء العملية
            if (extra && typeof extra.reply === 'function') {
                await extra.reply("جاري جلب أرقام الأعضاء وإنشاء ملف Excel... انتظر لحظة ⏳");
            }

            // 2. جلب بيانات المجموعة والأعضاء من نسخة الـ sock
            const groupMetadata = await sock.groupMetadata(groupJid);
            const participants = groupMetadata.participants;

            // 3. تجهيز البيانات وترتيبها داخل جدول الاكسل
            const rows = participants.map((participant, index) => {
                // استخراج الرقم بدون امتداد الواتساب @s.whatsapp.net
                const phoneNumber = participant.id.split('@')[0];
                
                // تحديد رتبة العضو (مالك، مشرف، عضو عادي)
                let role = 'عضو عادي';
                if (participant.admin === 'superadmin') {
                    role = 'منشئ المجموعة (المالك)';
                } else if (participant.admin === 'admin') {
                    role = 'مشرف (Admin)';
                }

                return {
                    "م": index + 1,
                    "رقم الهاتف": `+${phoneNumber}`,
                    "الرتبة في المجموعة": role
                };
            });

            // 4. إنشاء كتاب العمل (Workbook) وورقة العمل (Worksheet) الخاصة بالاكسل
            const worksheet = XLSX.utils.json_to_sheet(rows);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "قائمة الأعضاء");

            // 5. تحويل ملف الاكسل إلى "Buffer" في الذاكرة لإرساله مباشرة بدون الحاجة لحفظه على الجهاز
            const fileBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

            // اسم الملف الذي سيظهر للمستخدم
            const fileName = `أعضاء_${groupMetadata.subject || 'المجموعة'}.xlsx`;

            // 6. إرسال ملف الـ Excel للمجموعة
            await sock.sendMessage(groupJid, {
                document: fileBuffer,
                mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                fileName: fileName,
                caption: `📊 تم استخراج أرقام الأعضاء بنجاح!\n👥 إجمالي الأعضاء: ${participants.length}`
            }, { quoted: msg });

        } catch (error) {
            console.error("خطأ في أمر سحب الأعضاء:", error);
            if (extra && typeof extra.reply === 'function') {
                await extra.reply("❌ عذراً، حدث خطأ أثناء محاولة استخراج الأرقام وإنشاء الملف.");
            }
        }
    }
};
