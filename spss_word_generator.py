"""
SPSS Word Generator - Professional Algerian Thesis Format
=========================================================
Generates professional Word documents for SPSS analysis results
Formatted according to Algerian academic thesis standards

Supported analyses:
- Descriptive Statistics
- T-Test (Independent Samples)
- ANOVA (One-Way)
- Correlation (Pearson/Spearman)
- Regression (Multiple Linear)
- Chi-Square Test
- Cronbach's Alpha

Author: Automated SPSS Analysis System
Date: December 2024
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io


class SPSSWordGenerator:
    """Professional Word document generator for SPSS results"""
    
    def __init__(self):
        self.doc = Document()
        self._setup_document()
        
    def _setup_document(self):
        """Configure document with Algerian thesis standards"""
        # Page setup: A4 size
        section = self.doc.sections[0]
        section.page_height = Cm(29.7)  # A4 height
        section.page_width = Cm(21.0)   # A4 width
        
        # Margins (Algerian standard)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)   # Wider for binding
        section.right_margin = Cm(2.0)
        
        # Default font
        style = self.doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        
        # RTL support for Arabic
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    
    def _add_title(self, text, level=1):
        """Add formatted title"""
        title = self.doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run = title.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(16 if level == 1 else 14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        title.paragraph_format.space_after = Pt(12)
        return title
    
    def _add_section_header(self, text):
        """Add section header"""
        header = self.doc.add_paragraph()
        header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        run = header.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 139)  # Dark blue
        
        header.paragraph_format.space_before = Pt(12)
        header.paragraph_format.space_after = Pt(6)
        return header
    
    def _add_paragraph(self, text, align='right', bold=False):
        """Add formatted paragraph"""
        para = self.doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align == 'right' else WD_ALIGN_PARAGRAPH.LEFT
        
        run = para.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = bold
        
        return para
    
    def _create_table(self, rows, cols, headers=None):
        """Create formatted SPSS-style table"""
        table = self.doc.add_table(rows=rows, cols=cols)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Set borders (black, 0.5pt)
        for row in table.rows:
            for cell in row.cells:
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()
                
                # Add borders
                tcBorders = OxmlElement('w:tcBorders')
                for border_name in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{border_name}')
                    border.set(qn('w:val'), 'single')
                    border.set(qn('w:sz'), '6')  # 0.5pt
                    border.set(qn('w:color'), '000000')
                    tcBorders.append(border)
                
                tcPr.append(tcBorders)
        
        # Header row formatting
        if headers:
            header_cells = table.rows[0].cells
            for i, header_text in enumerate(headers):
                cell = header_cells[i]
                cell.text = header_text
                
                # Gray background
                shading = OxmlElement('w:shd')
                shading.set(qn('w:fill'), 'D9D9D9')
                cell._element.get_or_add_tcPr().append(shading)
                
                # Center align and bold
                paragraph = cell.paragraphs[0]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.runs[0]
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.name = 'Times New Roman'
        
        return table
    
    def _fill_table_cell(self, cell, text, align='center', bold=False):
        """Fill table cell with formatted text"""
        cell.text = str(text)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == 'center' else WD_ALIGN_PARAGRAPH.RIGHT
        
        run = paragraph.runs[0]
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.bold = bold
    
    def generate_descriptive(self, results):
        """Generate Descriptive Statistics report"""
        self._add_title("التحليل الإحصائي الوصفي\nDescriptive Statistics")
        self.doc.add_paragraph()
        
        # Introduction
        self._add_section_header("أولاً: الإحصاء الوصفي للمتغيرات")
        self._add_paragraph(
            "يعرض الجدول التالي الإحصاءات الوصفية للمتغيرات الرقمية المدروسة، "
            "حيث يتضمن عدد المشاهدات، المتوسط الحسابي، الانحراف المعياري، "
            "والقيم الدنيا والعليا لكل متغير."
        )
        self.doc.add_paragraph()
        
        # Numeric variables table
        if results.get('متغيرات_رقمية'):
            vars_data = results['متغيرات_رقمية']
            
            # Create table
            table = self._create_table(
                rows=len(vars_data) + 1,
                cols=6,
                headers=['المتغير', 'N', 'Mean', 'Std. Deviation', 'Minimum', 'Maximum']
            )
            
            # Fill data
            for i, var in enumerate(vars_data, start=1):
                cells = table.rows[i].cells
                self._fill_table_cell(cells[0], var['المتغير'], align='right')
                self._fill_table_cell(cells[1], var['العدد'])
                self._fill_table_cell(cells[2], f"{var['المتوسط']:.2f}")
                self._fill_table_cell(cells[3], f"{var['الانحراف_المعياري']:.2f}")
                self._fill_table_cell(cells[4], f"{var['أصغر_قيمة']:.2f}")
                self._fill_table_cell(cells[5], f"{var['أكبر_قيمة']:.2f}")
            
            self.doc.add_paragraph()
        
        # Interpretation
        self._add_section_header("ثانياً: التفسير الأكاديمي")
        self._add_paragraph(
            "تشير النتائج الإحصائية الوصفية إلى تباين في قيم المتغيرات المدروسة، "
            "حيث يمكن ملاحظة اختلاف المتوسطات الحسابية والانحرافات المعيارية. "
            "هذا التباين يعكس التنوع في استجابات أفراد العينة، ويساعد في فهم "
            "الخصائص العامة للبيانات قبل إجراء التحليلات الاستدلالية."
        )
        self.doc.add_paragraph()
        
        # Writing guidelines
        self._add_section_header("ثالثاً: كيفية الكتابة في المذكرة")
        self._add_paragraph(
            '▪ في فصل الإجراءات المنهجية:\n'
            '"تم استخدام الإحصاء الوصفي لتحليل خصائص العينة، حيث تم حساب '
            'المتوسطات الحسابية والانحرافات المعيارية للمتغيرات الرقمية."\n\n'
            '▪ في فصل النتائج:\n'
            'يُدرج الجدول أعلاه مع تفسير مختصر للنتائج البارزة.',
            align='right'
        )
        
        return self.doc
    
    def generate_ttest(self, results):
        """Generate Independent Samples T-Test report"""
        self._add_title("اختبار T للعينات المستقلة\nIndependent Samples T-Test")
        self.doc.add_paragraph()
        
        if 'error' in results:
            self._add_paragraph(f"❌ خطأ: {results['error']}")
            return self.doc
        
        # Introduction
        self._add_section_header("أولاً: إحصاءات المجموعات")
        self._add_paragraph(
            "يعرض الجدول التالي الإحصاءات الوصفية لكل مجموعة من مجموعتي المقارنة."
        )
        self.doc.add_paragraph()
        
        # Group Statistics Table
        table1 = self._create_table(
            rows=3,
            cols=4,
            headers=['المجموعة', 'N', 'Mean', 'Std. Deviation']
        )
        
        # Group 1
        cells = table1.rows[1].cells
        self._fill_table_cell(cells[0], results['المجموعة_1']['الاسم'], align='right')
        self._fill_table_cell(cells[1], results['المجموعة_1']['العدد'])
        self._fill_table_cell(cells[2], f"{results['المجموعة_1']['المتوسط']:.2f}")
        self._fill_table_cell(cells[3], f"{results['المجموعة_1']['الانحراف']:.2f}")
        
        # Group 2
        cells = table1.rows[2].cells
        self._fill_table_cell(cells[0], results['المجموعة_2']['الاسم'], align='right')
        self._fill_table_cell(cells[1], results['المجموعة_2']['العدد'])
        self._fill_table_cell(cells[2], f"{results['المجموعة_2']['المتوسط']:.2f}")
        self._fill_table_cell(cells[3], f"{results['المجموعة_2']['الانحراف']:.2f}")
        
        self.doc.add_paragraph()
        
        # T-Test Results
        self._add_section_header("ثانياً: نتائج اختبار T")
        self._add_paragraph(
            "يوضح الجدول التالي نتائج اختبار T للفروق بين المجموعتين."
        )
        self.doc.add_paragraph()
        
        table2 = self._create_table(
            rows=2,
            cols=4,
            headers=['t', 'df', 'Sig. (2-tailed)', "Cohen's d"]
        )
        
        cells = table2.rows[1].cells
        self._fill_table_cell(cells[0], f"{results['t']:.3f}")
        self._fill_table_cell(cells[1], results['df'])
        self._fill_table_cell(cells[2], f"{results['p']:.4f}")
        self._fill_table_cell(cells[3], f"{results['cohens_d']:.3f}")
        
        self.doc.add_paragraph()
        
        # Interpretation
        self._add_section_header("ثالثاً: التفسير الأكاديمي")
        
        if results['دال']:
            interp = (
                f"أظهرت نتائج اختبار T وجود فروق ذات دلالة إحصائية بين المجموعتين "
                f"عند مستوى دلالة {results['مستوى_الدلالة']}, حيث بلغت قيمة t = {results['t']:.3f} "
                f"بدرجات حرية df = {results['df']}, وقيمة p = {results['p']:.4f}. "
                f"وبلغ حجم الأثر (Cohen's d = {results['cohens_d']:.3f}) وهو {results['حجم_الأثر']}، "
                f"مما يشير إلى أن الفروق بين المجموعتين ذات أهمية عملية."
            )
        else:
            interp = (
                f"أظهرت نتائج اختبار T عدم وجود فروق ذات دلالة إحصائية بين المجموعتين "
                f"عند مستوى دلالة 0.05, حيث بلغت قيمة t = {results['t']:.3f} "
                f"بدرجات حرية df = {results['df']}, وقيمة p = {results['p']:.4f}، "
                f"وهي قيمة أكبر من 0.05، مما يعني قبول الفرضية الصفرية."
            )
        
        self._add_paragraph(interp)
        
        return self.doc
    
    def generate_anova(self, results):
        """Generate One-Way ANOVA report"""
        self._add_title("تحليل التباين الأحادي
One-Way ANOVA")
        self.doc.add_paragraph()
        
        if 'error' in results:
            self._add_paragraph(f"❌ خطأ: {results['error']}")
            return self.doc
        
        # ===== NEW: Methodological Info =====
        self._add_section_header("📋 معلومات التحليل:")
        self._add_paragraph(f"• الاختبار: تحليل التباين الأحادي (One-Way ANOVA)")
        if 'إحصاءات_المجموعات' in results:
            self._add_paragraph(f"• عدد المجموعات: {len(results['إحصاءات_المجموعات'])}")
        self._add_paragraph(f"• العدد الكلي: N = {results.get('N', 'غير محدد')}")
        self._add_paragraph(f"• مستوى الدلالة: α = 0.05")
        self.doc.add_paragraph()
        
        # ===== NEW: Group Descriptive Statistics =====
        self._add_section_header("📊 أولاً: الإحصاءات الوصفية للمجموعات")
        self._add_paragraph(
            "يعرض الجدول التالي الإحصاءات الوصفية لكل مجموعة من مجموعات المتغير المستقل."
        )
        self.doc.add_paragraph()
        
        if 'إحصاءات_المجموعات' in results:
            groups = results['إحصاءات_المجموعات']
            
            table = self._create_table(
                rows=len(groups) + 1,
                cols=4,
                headers=['المجموعة', 'N', 'Mean', 'Std. Deviation']
            )
            
            for i, (group_name, stats) in enumerate(groups.items(), start=1):
                cells = table.rows[i].cells
                self._fill_table_cell(cells[0], str(group_name), align='right', bold=True)
                self._fill_table_cell(cells[1], stats.get('العدد', '-'))
                self._fill_table_cell(cells[2], f"{stats.get('المتوسط', 0):.2f}")
                self._fill_table_cell(cells[3], f"{stats.get('الانحراف_المعياري', 0):.2f}")
            
            self.doc.add_paragraph()
        
        # ===== ANOVA Table =====
        self._add_section_header("📈 ثانياً: جدول تحليل التباين ANOVA")
        self._add_paragraph(
            "يعرض الجدول التالي نتائج تحليل التباين الأحادي للفروق بين المجموعات."
        )
        self.doc.add_paragraph()
        
        table = self._create_table(
            rows=4,
            cols=6,
            headers=['مصدر التباين', 'Sum of Squares', 'df', 'Mean Square', 'F', 'Sig.']
        )
        
        # Between Groups
        cells = table.rows[1].cells
        self._fill_table_cell(cells[0], 'بين المجموعات', align='right')
        self._fill_table_cell(cells[1], f"{results['بين_المجموعات']['مجموع_المربعات']:.3f}")
        self._fill_table_cell(cells[2], results['بين_المجموعات']['درجات_الحرية'])
        self._fill_table_cell(cells[3], f"{results['بين_المجموعات']['متوسط_المربعات']:.3f}")
        self._fill_table_cell(cells[4], f"{results['F']:.3f}")
        self._fill_table_cell(cells[5], f"{results['p']:.4f}")
        
        # Within Groups
        cells = table.rows[2].cells
        self._fill_table_cell(cells[0], 'داخل المجموعات', align='right')
        self._fill_table_cell(cells[1], f"{results['داخل_المجموعات']['مجموع_المربعات']:.3f}")
        self._fill_table_cell(cells[2], results['داخل_المجموعات']['درجات_الحرية'])
        self._fill_table_cell(cells[3], f"{results['داخل_المجموعات']['متوسط_المربعات']:.3f}")
        self._fill_table_cell(cells[4], '-')
        self._fill_table_cell(cells[5], '-')
        
        # Total
        cells = table.rows[3].cells
        self._fill_table_cell(cells[0], 'المجموع', align='right')
        self._fill_table_cell(cells[1], f"{results['الكلي']['مجموع_المربعات']:.3f}")
        self._fill_table_cell(cells[2], results['الكلي']['درجات_الحرية'])
        self._fill_table_cell(cells[3], '-')
        self._fill_table_cell(cells[4], '-')
        self._fill_table_cell(cells[5], '-')
        
        self.doc.add_paragraph()
        
        # Interpretation
        self._add_section_header("📖 ثالثاً: التفسير الأكاديمي")
        
        if results['دال']:
            interp = (
                f"أظهرت نتائج تحليل التباين الأحادي (ANOVA) وجود فروق ذات دلالة إحصائية "
                f"بين المجموعات عند مستوى دلالة {results['مستوى_الدلالة']}, حيث بلغت "
                f"قيمة F = {results['F']:.3f} بدرجات حرية "
                f"({results['بين_المجموعات']['درجات_الحرية']}, {results['داخل_المجموعات']['درجات_الحرية']}), "
                f"وقيمة p = {results['p']:.4f}. وبلغ حجم الأثر (Eta Squared = {results['eta_squared']:.3f}) "
                f"وهو {results['حجم_الأثر']}، مما يشير إلى وجود فروق جوهرية بين المجموعات."
            )
        else:
            interp = (
                f"أظهرت نتائج تحليل التباين الأحادي (ANOVA) عدم وجود فروق ذات دلالة إحصائية "
                f"بين المجموعات عند مستوى دلالة 0.05, حيث بلغت قيمة F = {results['F']:.3f} "
                f"بدرجات حرية ({results['بين_المجموعات']['درجات_الحرية']}, "
                f"{results['داخل_المجموعات']['درجات_الحرية']}), وقيمة p = {results['p']:.4f}، "
                f"وهي قيمة أكبر من 0.05."
            )
        
        self._add_paragraph(interp)
        
        # ===== NEW: Writing Guide =====
        self.doc.add_paragraph()
        self._add_section_header("📝 رابعاً: كيفية الكتابة في المذكرة")
        
        self._add_paragraph("• في فصل الإجراءات المنهجية:", bold=True)
        self._add_paragraph(
            '"تم استخدام اختبار تحليل التباين الأحادي (One-Way ANOVA) للكشف عن الفروق بين '
            'المجموعات، حيث بلغت العينة الكلية N = ' + str(results.get('N', 'X')) + '."'
        )
        
        self.doc.add_paragraph()
        self._add_paragraph("• في فصل النتائج:", bold=True)
        self._add_paragraph(
            '"أظهرت نتائج تحليل التباين الأحادي وجود فروق دالة إحصائياً بين المجموعات '
            '(F = X.XX, p < 0.05), مما يدل على تأثير [المتغير المستقل] على [المتغير التابع]."'
        )
        
        return self.doc

    def generate_correlation(self, results):
        """Generate Correlation Analysis report"""
        self._add_title("تحليل الارتباط
Correlation Analysis")
        self.doc.add_paragraph()
        
        if 'error' in results:
            self._add_paragraph(f"❌ خطأ: {results['error']}")
            return self.doc
        
        # ===== NEW: Methodological Info =====
        self._add_section_header("📋 معلومات التحليل:")
        method_ar = "بيرسون" if results.get('method') == 'pearson' else "سبيرمان"
        method_en = "Pearson" if results.get('method') == 'pearson' else "Spearman"
        self._add_paragraph(f"• الاختبار: معامل ارتباط {method_ar} ({method_en} Correlation)")
        self._add_paragraph(f"• العدد الكلي: N = {results.get('N', 'غير محدد')}")
        self._add_paragraph(f"• مستوى الدلالة: α = 0.05")
        self.doc.add_paragraph()
        
        # ===== NEW: Descriptive Statistics =====
        self._add_section_header("📊 أولاً: الإحصاءات الوصفية للمتغيرات")
        self._add_paragraph(
            "يعرض الجدول التالي الإحصاءات الوصفية للمتغيرات المدروسة في تحليل الارتباط."
        )
        self.doc.add_paragraph()
        
        if 'إحصاءات_وصفية' in results:
            descriptives = results['إحصاءات_وصفية']
            
            table = self._create_table(
                rows=len(descriptives) + 1,
                cols=4,
                headers=['المتغير', 'N', 'Mean', 'Std. Deviation']
            )
            
            for i, (var_name, stats) in enumerate(descriptives.items(), start=1):
                cells = table.rows[i].cells
                self._fill_table_cell(cells[0], str(var_name), align='right', bold=True)
                self._fill_table_cell(cells[1], stats.get('N', '-'))
                self._fill_table_cell(cells[2], f"{stats.get('Mean', 0):.2f}")
                self._fill_table_cell(cells[3], f"{stats.get('SD', 0):.2f}")
            
            self.doc.add_paragraph()
        
        # ===== Correlation Matrix =====
        self._add_section_header("📈 ثانياً: مصفوفة الارتباط")
        self._add_paragraph(
            "يعرض الجدول التالي معاملات الارتباط بين المتغيرات، حيث تشير النجوم إلى مستوى الدلالة الإحصائية "
            "(* p < 0.05, ** p < 0.01, *** p < 0.001)."
        )
        self.doc.add_paragraph()
        
        if 'مصفوفة_الارتباط' in results:
            matrix = results['مصفوفة_الارتباط']
            variables = list(matrix.keys())
            
            # Create table
            table = self._create_table(
                rows=len(variables) + 1,
                cols=len(variables) + 1,
                headers=[''] + variables
            )
            
            # Fill matrix
            for i, var1 in enumerate(variables, start=1):
                cells = table.rows[i].cells
                self._fill_table_cell(cells[0], var1, align='right', bold=True)
                
                for j, var2 in enumerate(variables, start=1):
                    r_value = matrix[var1][var2]['r']
                    p_value = matrix[var1][var2]['p']
                    
                    # Format with significance stars
                    if p_value < 0.001:
                        sig_text = f"{r_value:.3f}***"
                    elif p_value < 0.01:
                        sig_text = f"{r_value:.3f}**"
                    elif p_value < 0.05:
                        sig_text = f"{r_value:.3f}*"
                    else:
                        sig_text = f"{r_value:.3f}"
                    
                    self._fill_table_cell(cells[j], sig_text)
            
            self.doc.add_paragraph()
            
            # ===== NEW: Note about N =====
            note = self.doc.add_paragraph()
            note.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = note.add_run(f"Note: N = {results.get('N', 'X')} for all correlations.")
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)
            run.font.italic = True
            self.doc.add_paragraph()
        
        # Interpretation
        self._add_section_header("📖 ثالثاً: التفسير الأكاديمي")
        
        if 'نتائج_دالة' in results and results['نتائج_دالة']:
            interp = "أظهرت نتائج تحليل الارتباط وجود علاقات ذات دلالة إحصائية بين بعض المتغيرات:

"
            
            for result in results['نتائج_دالة']:
                direction = "موجبة" if result['r'] > 0 else "سالبة"
                strength = result.get('قوة', 'متوسطة')
                interp += (
                    f"• العلاقة بين {result['var1']} و {result['var2']}: "
                    f"علاقة {direction} {strength} (r = {result['r']:.3f}, p = {result['p']:.4f})
"
                )
        else:
            interp = "أظهرت نتائج تحليل الارتباط عدم وجود علاقات ذات دلالة إحصائية بين المتغيرات عند مستوى دلالة 0.05."
        
        self._add_paragraph(interp)
        
        # ===== NEW: Writing Guide =====
        self.doc.add_paragraph()
        self._add_section_header("📝 رابعاً: كيفية الكتابة في المذكرة")
        
        self._add_paragraph("• في فصل الإجراءات المنهجية:", bold=True)
        method_ar = "بيرسون" if results.get('method') == 'pearson' else "سبيرمان"
        self._add_paragraph(
            f'"تم استخدام معامل ارتباط {method_ar} لقياس العلاقة بين المتغيرات، '
            f'حيث بلغت العينة N = {results.get("N", "X")}."'
        )
        
        self.doc.add_paragraph()
        self._add_paragraph("• في فصل النتائج:", bold=True)
        self._add_paragraph(
            '"أظهرت نتائج تحليل الارتباط وجود علاقة [موجبة/سالبة] [ضعيفة/متوسطة/قوية] '
            'ذات دلالة إحصائية بين [المتغير الأول] و[المتغير الثاني] (r = X.XX, p < 0.05)."'
        )
        
        return self.doc

    def generate_regression(self, results):
        """Generate Multiple Regression Analysis report"""
        self._add_title("تحليل الانحدار المتعدد\nMultiple Regression Analysis")
        self.doc.add_paragraph()
        
        if 'error' in results:
            self._add_paragraph(f"❌ خطأ: {results['error']}")
            return self.doc
        
        # Model Summary
        self._add_section_header("أولاً: ملخص النموذج - Model Summary")
        self._add_paragraph("يوضح الجدول التالي جودة النموذج الإحصائي.")
        self.doc.add_paragraph()
        
        table1 = self._create_table(
            rows=2,
            cols=4,
            headers=['R', 'R Square', 'Adjusted R Square', 'Std. Error']
        )
        
        cells = table1.rows[1].cells
        self._fill_table_cell(cells[0], f"{results['R']:.3f}")
        self._fill_table_cell(cells[1], f"{results['R2']:.3f}")
        self._fill_table_cell(cells[2], f"{results['R2_المعدل']:.3f}")
        self._fill_table_cell(cells[3], f"{results.get('الخطأ_المعياري', 0):.3f}")
        
        self.doc.add_paragraph()
        
        # ANOVA Table
        self._add_section_header("ثانياً: جدول تحليل التباين ANOVA")
        self._add_paragraph("يوضح الجدول التالي مدى معنوية النموذج ككل.")
        self.doc.add_paragraph()
        
        table2 = self._create_table(
            rows=3,
            cols=6,
            headers=['مصدر التباين', 'Sum of Squares', 'df', 'Mean Square', 'F', 'Sig.']
        )
        
        # Regression row
        cells = table2.rows[1].cells
        self._fill_table_cell(cells[0], 'الانحدار', align='right')
        self._fill_table_cell(cells[1], '-')
        self._fill_table_cell(cells[2], '-')
        self._fill_table_cell(cells[3], '-')
        self._fill_table_cell(cells[4], f"{results['F']:.3f}")
        self._fill_table_cell(cells[5], f"{results['p_model']:.4f}")
        
        # Residual row
        cells = table2.rows[2].cells
        self._fill_table_cell(cells[0], 'البواقي', align='right')
        self._fill_table_cell(cells[1], '-')
        self._fill_table_cell(cells[2], '-')
        self._fill_table_cell(cells[3], '-')
        self._fill_table_cell(cells[4], '-')
        self._fill_table_cell(cells[5], '-')
        
        self.doc.add_paragraph()
        
        # Coefficients Table
        self._add_section_header("ثالثاً: معاملات الانحدار - Coefficients")
        self._add_paragraph("يوضح الجدول التالي معاملات الانحدار لكل متغير مستقل.")
        self.doc.add_paragraph()
        
        if 'معاملات' in results:
            coefs = results['معاملات']
            table3 = self._create_table(
                rows=len(coefs) + 1,
                cols=5,
                headers=['المتغير', 'B', 'Std. Error', 't', 'Sig.']
            )
            
            for i, coef in enumerate(coefs, start=1):
                cells = table3.rows[i].cells
                self._fill_table_cell(cells[0], coef['المتغير'], align='right')
                self._fill_table_cell(cells[1], f"{coef['المعامل']:.3f}")
                self._fill_table_cell(cells[2], f"{coef.get('الخطأ_المعياري', 0):.3f}")
                self._fill_table_cell(cells[3], f"{coef.get('t', 0):.3f}")
                self._fill_table_cell(cells[4], f"{coef['p']:.4f}")
            
            self.doc.add_paragraph()
        
        # Interpretation
        self._add_section_header("رابعاً: التفسير الأكاديمي")
        interp = (
            f"أظهرت نتائج تحليل الانحدار المتعدد أن النموذج يفسر {results['R2']*100:.1f}% "
            f"من التباين في المتغير التابع (R² = {results['R2']:.3f}). "
        )
        
        if results['دال']:
            interp += (
                f"وأظهر اختبار F معنوية النموذج ككل (F = {results['F']:.3f}, "
                f"p = {results['p_model']:.4f}), مما يشير إلى أن المتغيرات المستقلة "
                f"مجتمعة لها تأثير دال إحصائياً على المتغير التابع."
            )
        else:
            interp += "إلا أن النموذج ككل غير دال إحصائياً عند مستوى 0.05."
        
        self._add_paragraph(interp)
        
        return self.doc
    
    def generate_chisquare(self, results):
        """Generate Chi-Square Test report"""
        self._add_title("اختبار مربع كاي
Chi-Square Test")
        self.doc.add_paragraph()
        
        if 'error' in results:
            self._add_paragraph(f"❌ خطأ: {results['error']}")
            return self.doc
        
        # ===== NEW: Methodological Info =====
        self._add_section_header("📋 معلومات التحليل:")
        self._add_paragraph(f"• الاختبار: اختبار مربع كاي للاستقلالية (Chi-Square Test of Independence)")
        self._add_paragraph(f"• المتغير الأول: {results.get('var1', 'غير محدد')}")
        self._add_paragraph(f"• المتغير الثاني: {results.get('var2', 'غير محدد')}")
        self._add_paragraph(f"• العدد الكلي: N = {results.get('N', 'غير محدد')}")
        self._add_paragraph(f"• مستوى الدلالة: α = 0.05")
        self.doc.add_paragraph()
        
        # ===== NEW: Crosstabulation Table =====
        self._add_section_header("📊 أولاً: جدول التوافق (Crosstabulation)")
        self._add_paragraph(
            "يعرض الجدول التالي التوزيع التكراري للحالات حسب فئات المتغيرين المدروسين."
        )
        self.doc.add_paragraph()
        
        if 'جدول_التوافق' in results:
            crosstab = results['جدول_التوافق']
            
            # Get categories
            row_categories = list(crosstab.keys())
            col_categories = list(crosstab[row_categories[0]].keys())
            
            # Create table (rows + header + total row)
            table = self._create_table(
                rows=len(row_categories) + 2,
                cols=len(col_categories) + 2,
                headers=[''] + col_categories + ['المجموع']
            )
            
            # Calculate column totals
            col_totals = {col: 0 for col in col_categories}
            grand_total = 0
            
            # Fill data rows
            for i, row_cat in enumerate(row_categories, start=1):
                cells = table.rows[i].cells
                self._fill_table_cell(cells[0], str(row_cat), align='right', bold=True)
                
                row_total = 0
                for j, col_cat in enumerate(col_categories, start=1):
                    count = crosstab[row_cat][col_cat]
                    self._fill_table_cell(cells[j], str(count))
                    row_total += count
                    col_totals[col_cat] += count
                
                # Row total
                self._fill_table_cell(cells[-1], str(row_total), bold=True)
                grand_total += row_total
            
            # Total row
            last_row_cells = table.rows[-1].cells
            self._fill_table_cell(last_row_cells[0], 'المجموع', align='right', bold=True)
            
            for j, col_cat in enumerate(col_categories, start=1):
                self._fill_table_cell(last_row_cells[j], str(col_totals[col_cat]), bold=True)
            
            self._fill_table_cell(last_row_cells[-1], str(grand_total), bold=True)
            
            self.doc.add_paragraph()
        
        # ===== Chi-Square Results =====
        self._add_section_header("📈 ثانياً: نتائج اختبار مربع كاي")
        self._add_paragraph(
            "يعرض الجدول التالي نتائج اختبار مربع كاي للاستقلالية بين المتغيرين."
        )
        self.doc.add_paragraph()
        
        table = self._create_table(
            rows=2,
            cols=4,
            headers=['Chi-Square (χ²)', 'df', 'Asymp. Sig.', "Cramér's V"]
        )
        
        cells = table.rows[1].cells
        self._fill_table_cell(cells[0], f"{results['chi_square']:.3f}")
        self._fill_table_cell(cells[1], results['df'])
        self._fill_table_cell(cells[2], f"{results['p']:.4f}")
        
        if 'cramers_v' in results:
            self._fill_table_cell(cells[3], f"{results['cramers_v']:.3f}")
        else:
            self._fill_table_cell(cells[3], '-')
        
        self.doc.add_paragraph()
        
        # Interpretation
        self._add_section_header("📖 ثالثاً: التفسير الأكاديمي")
        
        if results.get('دال'):
            interp = (
                f"أظهرت نتائج اختبار مربع كاي وجود علاقة ذات دلالة إحصائية بين المتغيرين "
                f"عند مستوى دلالة {results.get('مستوى_الدلالة', '0.05')}, حيث بلغت قيمة "
                f"χ² = {results['chi_square']:.3f} بدرجات حرية df = {results['df']}, "
                f"وقيمة p = {results['p']:.4f}. "
            )
            
            if 'cramers_v' in results:
                strength = results.get('قوة_العلاقة', 'متوسطة')
                interp += (
                    f"وبلغ معامل كرامر (Cramér's V = {results['cramers_v']:.3f}) "
                    f"وهو يشير إلى علاقة {strength} بين المتغيرين."
                )
        else:
            interp = (
                f"أظهرت نتائج اختبار مربع كاي عدم وجود علاقة ذات دلالة إحصائية بين المتغيرين "
                f"عند مستوى دلالة 0.05, حيث بلغت قيمة χ² = {results['chi_square']:.3f} "
                f"بدرجات حرية df = {results['df']}, وقيمة p = {results['p']:.4f}، "
                f"وهي قيمة أكبر من 0.05، مما يدل على استقلالية المتغيرين."
            )
        
        self._add_paragraph(interp)
        
        # ===== NEW: Writing Guide =====
        self.doc.add_paragraph()
        self._add_section_header("📝 رابعاً: كيفية الكتابة في المذكرة")
        
        self._add_paragraph("• في فصل الإجراءات المنهجية:", bold=True)
        self._add_paragraph(
            '"تم استخدام اختبار مربع كاي (Chi-Square) للكشف عن العلاقة بين المتغيرين الاسميين، '
            'حيث بلغت العينة الكلية N = ' + str(results.get('N', 'X')) + '."'
        )
        
        self.doc.add_paragraph()
        self._add_paragraph("• في فصل النتائج:", bold=True)
        self._add_paragraph(
            '"أظهرت نتائج اختبار مربع كاي وجود علاقة دالة إحصائياً بين [المتغير الأول] '
            'و[المتغير الثاني] (χ² = X.XX, p < 0.05), مما يدل على وجود ارتباط بين المتغيرين."'
        )
        
        return self.doc

    def generate_cronbach(self, results):
        """Generate Cronbach's Alpha Reliability report"""
        self._add_title("معامل ألفا كرونباخ للثبات\nCronbach's Alpha Reliability")
        self.doc.add_paragraph()
        
        if 'error' in results:
            self._add_paragraph(f"❌ خطأ: {results['error']}")
            return self.doc
        
        # Introduction
        self._add_section_header("أولاً: إحصاءات الثبات - Reliability Statistics")
        self._add_paragraph(
            "يعرض الجدول التالي معامل ألفا كرونباخ الذي يقيس الاتساق الداخلي للمقياس."
        )
        self.doc.add_paragraph()
        
        # Reliability Statistics Table
        table1 = self._create_table(
            rows=2,
            cols=2,
            headers=["Cronbach's Alpha", 'N of Items']
        )
        
        cells = table1.rows[1].cells
        self._fill_table_cell(cells[0], f"{results['alpha']:.3f}")
        self._fill_table_cell(cells[1], results['عدد_البنود'])
        
        self.doc.add_paragraph()
        
        # Item Statistics
        self._add_section_header("ثانياً: إحصاءات البنود - Item Statistics")
        self._add_paragraph("يوضح الجدول التالي الإحصاءات الوصفية لكل بند في المقياس.")
        self.doc.add_paragraph()
        
        if 'إحصاءات_البنود' in results:
            items = results['إحصاءات_البنود']
            table2 = self._create_table(
                rows=len(items) + 1,
                cols=4,
                headers=['البند', 'Mean', 'Std. Deviation', 'Alpha if Deleted']
            )
            
            for i, item in enumerate(items, start=1):
                cells = table2.rows[i].cells
                self._fill_table_cell(cells[0], item['البند'], align='right')
                self._fill_table_cell(cells[1], f"{item['المتوسط']:.2f}")
                self._fill_table_cell(cells[2], f"{item['الانحراف']:.2f}")
                alpha_del = item.get('ألفا_إذا_حُذف')
                self._fill_table_cell(cells[3], f"{alpha_del:.3f}" if alpha_del else 'N/A')
            
            self.doc.add_paragraph()
        
        # Interpretation
        self._add_section_header("ثالثاً: التفسير الأكاديمي")
        
        interp = (
            f"بلغت قيمة معامل ألفا كرونباخ ({results['alpha']:.3f})، وهي قيمة تُصنف "
            f"على أنها {results['التصنيف']} وفقاً للمعايير المتعارف عليها. "
        )
        
        if results['alpha'] >= 0.70:
            interp += (
                "وهذا يشير إلى أن المقياس يتمتع بثبات داخلي جيد، مما يعني أن البنود "
                "متسقة فيما بينها وتقيس نفس البُنية النظرية. من خلال عمود 'Alpha if Deleted'، "
                "يمكن ملاحظة البنود التي قد يؤدي حذفها إلى تحسين أو خفض الثبات الكلي للمقياس."
            )
        else:
            interp += (
                "وهذا يشير إلى ضرورة مراجعة بعض البنود أو إضافة بنود جديدة لتحسين "
                "الثبات الداخلي للمقياس. يُوصى بفحص البنود ذات الارتباط المنخفض أو "
                "التي يؤدي حذفها إلى رفع قيمة ألفا."
            )
        
        self._add_paragraph(interp)
        self.doc.add_paragraph()
        
        # Writing Guidelines
        self._add_section_header("رابعاً: كيفية الكتابة في المذكرة")
        self._add_paragraph(
            f'▪ في فصل الإجراءات المنهجية:\n'
            f'"تم التحقق من ثبات المقياس باستخدام معامل ألفا كرونباخ، '
            f'حيث بلغت قيمته (α = {results["alpha"]:.3f})، وهي قيمة {results["التصنيف"]} '
            f'تشير إلى ثبات داخلي مناسب للمقياس."\n\n'
            f'▪ في جدول خصائص أدوات الدراسة:\n'
            f'يمكن إدراج جدول إحصاءات البنود أعلاه مباشرة.',
            align='right'
        )
        
        return self.doc
    
    def save_to_bytes(self):
        """Save document to bytes (for HTTP response)"""
        file_stream = io.BytesIO()
        self.doc.save(file_stream)
        file_stream.seek(0)
        return file_stream
    
    def save(self, filename):
        """Save document to file"""
        self.doc.save(filename)


# Main function for testing
if __name__ == "__main__":
    # Test example - Cronbach's Alpha
    generator = SPSSWordGenerator()
    
    test_results = {
        'alpha': 0.876,
        'عدد_البنود': 5,
        'حجم_العينة': 120,
        'التصنيف': 'ممتاز (Excellent)',
        'إحصاءات_البنود': [
            {'البند': 'البند 1', 'المتوسط': 3.45, 'الانحراف': 0.89, 'الارتباط_مع_المجموع': 0.67, 'ألفا_إذا_حُذف': 0.851},
            {'البند': 'البند 2', 'المتوسط': 3.78, 'الانحراف': 0.76, 'الارتباط_مع_المجموع': 0.72, 'ألفا_إذا_حُذف': 0.843},
            {'البند': 'البند 3', 'المتوسط': 3.56, 'الانحراف': 0.92, 'الارتباط_مع_المجموع': 0.68, 'ألفا_إذا_حُذف': 0.849},
            {'البند': 'البند 4', 'المتوسط': 3.92, 'الانحراف': 0.81, 'الارتباط_مع_المجموع': 0.75, 'ألفا_إذا_حُذف': 0.836},
            {'البند': 'البند 5', 'المتوسط': 3.67, 'الانحراف': 0.85, 'الارتباط_مع_المجموع': 0.70, 'ألفا_إذا_حُذف': 0.845},
        ]
    }
    
    generator.generate_cronbach(test_results)
    generator.save('/mnt/user-data/outputs/test_cronbach_report.docx')
    print("✅ Test Cronbach report generated successfully!")
