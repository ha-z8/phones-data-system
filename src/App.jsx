import React, { useState, useEffect } from 'react';
import { createClient } from '@supabase/supabase-js';
import { Document, Packer, Paragraph, TextRun, Table, TableCell, TableRow, WidthType, AlignmentType, BorderStyle, ShadingType } from 'docx';
import { saveAs } from 'file-saver';

const supabaseUrl = 'https://pnixqykxqoitsvqhbcsw.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBuaXhxeWt4cW9pdHN2cWhiY3N3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTIyMDgyNDUsImV4cCI6MjEwNzc4NDI0NX0.K7q9KcCjOwRxaSdomQoYVQp9sapbnI5_YS2MJiz93DM';
const supabase = createClient(supabaseUrl, supabaseKey);

export default function App() {
  const [students, setStudents] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [loading, setLoading] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);
  const [adminPassword, setAdminPassword] = useState('');

  useEffect(() => {
    fetchStudents();
  }, []);

  const fetchStudents = async () => {
    setLoading(true);
    const { data, error } = await supabase
      .from('phones_data')
      .select('*')
      .order('id', { ascending: true });
    
    if (error) {
      console.error('خطأ في جلب البيانات:', error);
      alert('فشل في جلب البيانات من Supabase. تأكد من تنفيذ أمر SQL لإيقاف الـ RLS.');
    } else {
      setStudents(data || []);
    }
    setLoading(false);
  };

  const handlePhoneUpdate = async (id, newPhoneType) => {
    const { error } = await supabase
      .from('phones_data')
      .update({ phone_type: newPhoneType })
      .eq('id', id);

    if (error) {
      alert('حدث خطأ أثناء حفظ التحديث في قاعدة البيانات');
    } else {
      setStudents(students.map(s => s.id === id ? { ...s, phone_type: newPhoneType } : s));
    }
  };

  const handleAdminLogin = (e) => {
    e.preventDefault();
    if (adminPassword === 'Aa12345678') {
      setIsAdmin(true);
    } else {
      alert('كلمة المرور غير صحيحة');
    }
  };

  const generateWordDocument = async () => {
    try {
      if (students.length === 0) {
        alert('لا توجد بيانات طلاب لتصديرها!');
        return;
      }

      const cabinetsMap = {};

      students.forEach((student) => {
        let rawCab = String(student.cabinet_number || '1');
        const cleanCab = rawCab.replace(/\D/g, '') || rawCab.trim();

        if (!cabinetsMap[cleanCab]) {
          cabinetsMap[cleanCab] = [];
        }
        cabinetsMap[cleanCab].push(student);
      });

      const docSections = [];

      const tableBorders = {
        top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      };

      const colWidths = [
        { size: 23, type: WidthType.PERCENTAGE }, // نوع الجوال
        { size: 22, type: WidthType.PERCENTAGE }, // القسم
        { size: 45, type: WidthType.PERCENTAGE }, // الاسم
        { size: 10, type: WidthType.PERCENTAGE }, // ت
      ];

      const sortedCabinets = Object.keys(cabinetsMap).sort((a, b) => {
        return (Number(a) || 0) - (Number(b) || 0);
      });

      sortedCabinets.forEach((cabName) => {
        const cabStudents = cabinetsMap[cabName];

        docSections.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 150 },
            bidirectional: true,
            children: [
              new TextRun({
                text: `دولاب رقم ( ${cabName} ) متوسط`,
                bold: true,
                size: 28,
                underline: { type: "single" },
                font: "Arial"
              })
            ]
          })
        );

        const tableRows = [
          new TableRow({
            tableHeader: true,
            cantSplit: true,
            children: [
              new TableCell({
                width: colWidths[0],
                shading: { fill: "0047AB", type: ShadingType.CLEAR, color: "auto" },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    bidirectional: true,
                    children: [new TextRun({ text: "نوع الجوال", bold: true, color: "FFFFFF", size: 24, font: "Arial" })]
                  })
                ],
              }),
              new TableCell({
                width: colWidths[1],
                shading: { fill: "0047AB", type: ShadingType.CLEAR, color: "auto" },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    bidirectional: true,
                    children: [new TextRun({ text: "القسم", bold: true, color: "FFFFFF", size: 24, font: "Arial" })]
                  })
                ],
              }),
              new TableCell({
                width: colWidths[2],
                shading: { fill: "0047AB", type: ShadingType.CLEAR, color: "auto" },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    bidirectional: true,
                    children: [new TextRun({ text: "الاسم", bold: true, color: "FFFFFF", size: 24, font: "Arial" })]
                  })
                ],
              }),
              new TableCell({
                width: colWidths[3],
                shading: { fill: "0047AB", type: ShadingType.CLEAR, color: "auto" },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    bidirectional: true,
                    children: [new TextRun({ text: "ت", bold: true, color: "FFFFFF", size: 24, font: "Arial" })]
                  })
                ],
              }),
            ],
          }),
        ];

        cabStudents.forEach((student, index) => {
          tableRows.push(
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: colWidths[0],
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      bidirectional: true,
                      children: [new TextRun({ text: student.phone_type || '', size: 22, font: "Arial" })]
                    })
                  ],
                }),
                new TableCell({
                  width: colWidths[1],
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      bidirectional: true,
                      children: [new TextRun({ text: student.section || 'متوسط', color: "0047AB", size: 22, font: "Arial" })]
                    })
                  ],
                }),
                new TableCell({
                  width: colWidths[2],
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      bidirectional: true,
                      children: [new TextRun({ text: student.student_name || '', size: 22, font: "Arial" })]
                    })
                  ],
                }),
                new TableCell({
                  width: colWidths[3],
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      bidirectional: true,
                      children: [new TextRun({ text: String(index + 1), size: 22, font: "Arial" })]
                    })
                  ],
                }),
              ],
            })
          );
        });

        const table = new Table({
          rows: tableRows,
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          cantSplit: true,
          alignment: AlignmentType.CENTER,
        });

        docSections.push(table);
        docSections.push(new Paragraph({ text: "", spacing: { after: 350 } }));
      });

      const doc = new Document({
        sections: [{ 
          properties: { 
            page: { 
              bidi: true,
              margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 }
            } 
          },
          children: docSections 
        }],
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, 'جدول_جوالات_الطلاب_المتوسط.docx');
    } catch (error) {
      console.error('خطأ في توليد ملف الورد:', error);
      alert('حدث خطأ أثناء تصدير ملف الورد.');
    }
  };

  const filteredStudents = students.filter(s => 
    s.student_name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    String(s.cabinet_number).toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="min-h-screen bg-emerald-50/50 py-4 px-2 sm:p-6 font-sans" dir="rtl">
      <div className="max-w-6xl mx-auto bg-white shadow-xl rounded-2xl p-4 sm:p-6 border border-emerald-100">
        
        {/* الهيدر والعنوان متجاوب للشاشات الصغيرة */}
        <header className="text-center mb-6 sm:mb-8 border-b border-emerald-100 pb-4 sm:pb-6">
          <h1 className="text-xl sm:text-3xl font-extrabold text-emerald-900 leading-snug">
            نظام تسجيل بيانات جوالات الطلاب
          </h1>
          <p className="text-emerald-700 mt-2 text-xs sm:text-sm">
            ابحث عن اسمك أو رقم الدولاب، أدخل نوع جوالك، وسيتم الحفظ المباشر
          </p>
        </header>

        {/* قسم البحث وتسجيل الدخول المتجاوب */}
        <div className="mb-6 flex flex-col gap-3 sm:gap-4 justify-between items-center bg-emerald-50/80 p-3 sm:p-4 rounded-xl border border-emerald-200">
          <input
            type="text"
            placeholder="🔍 ابحث باسم الطالب أو رقم الدولاب..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full p-3 border border-emerald-300 rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500 text-sm shadow-sm"
          />
          
          <div className="w-full flex justify-end">
            {!isAdmin ? (
              <form onSubmit={handleAdminLogin} className="flex flex-col sm:flex-row gap-2 w-full">
                <input
                  type="password"
                  placeholder="كلمة مرور المسؤول"
                  value={adminPassword}
                  onChange={(e) => setAdminPassword(e.target.value)}
                  className="p-3 border border-emerald-300 rounded-lg text-sm w-full bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500"
                />
                <button type="submit" className="bg-emerald-800 text-white px-4 py-3 rounded-lg text-sm hover:bg-emerald-900 transition font-medium whitespace-nowrap shadow-sm w-full sm:w-auto">
                  دخول المسؤول
                </button>
              </form>
            ) : (
              <button
                onClick={generateWordDocument}
                className="bg-emerald-600 text-white px-5 py-3 rounded-lg font-semibold hover:bg-emerald-700 transition shadow-md flex items-center gap-2 w-full justify-center text-sm"
              >
                📥 تحميل ملف الورد (جاهز ويعمل بامتياز)
              </button>
            )}
          </div>
        </div>

        {loading ? (
          <div className="text-center py-16">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-4 border-emerald-600 border-t-transparent"></div>
            <p className="text-emerald-800 mt-3 text-sm font-medium">جاري تحميل البيانات...</p>
          </div>
        ) : (
          /* جدول متجاوب يدعم الأجهزة الصغيرة مع ألوان متناوبة للصفوف الكاملة وتأثير Hover */
          <div className="overflow-x-auto shadow-sm rounded-xl border border-emerald-200">
            <table className="w-full min-w-[600px] border-collapse text-right bg-white text-sm">
              <thead>
                <tr className="bg-emerald-800 text-white text-xs uppercase tracking-wider">
                  <th className="p-3 border-l border-emerald-700 w-16 text-center">الدولاب</th>
                  <th className="p-3 border-l border-emerald-700">اسم الطالب</th>
                  <th className="p-3 border-l border-emerald-700 w-28 text-center">القسم</th>
                  <th className="p-3">نوع الجوال (اكتب الجهاز وسيتم الحفظ تلقائياً)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-emerald-100">
                {filteredStudents.length > 0 ? (
                  filteredStudents.map((student, idx) => (
                    <tr 
                      key={student.id} 
                      className={`transition-colors duration-150 hover:bg-emerald-100/70 ${
                        idx % 2 === 0 ? 'bg-white' : 'bg-emerald-50/50'
                      }`}
                    >
                      <td className="p-3 text-emerald-900 font-medium text-center whitespace-nowrap">{student.cabinet_number}</td>
                      <td className="p-3 font-semibold text-gray-900">{student.student_name}</td>
                      <td className="p-3 text-emerald-700 text-center whitespace-nowrap">{student.section}</td>
                      <td className="p-3">
                        <input
                          type="text"
                          defaultValue={student.phone_type || ''}
                          onBlur={(e) => handlePhoneUpdate(student.id, e.target.value)}
                          placeholder="اكتب نوع الجوال هنا..."
                          className="w-full p-2 sm:p-2.5 border border-emerald-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500 text-sm bg-white shadow-inner transition"
                        />
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan="4" className="text-center py-10 text-emerald-800 bg-emerald-50/50">
                      لا توجد نتائج مطابقة للبحث أو أن قاعدة البيانات فارغة
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}

      </div>
    </div>
  );
}