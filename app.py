"""
نظام التحليل الإحصائي الآلي لمذكرات التخرج - الجزائر
الإصدار: 1.0
التاريخ: 2024
"""

from flask import Flask, request, jsonify
import pandas as pd
import numpy as np
from scipy import stats
import statsmodels.api as sm
from datetime import datetime
import requests
from io import BytesIO
import re
import traceback

app = Flask(__name__)

# إعدادات الأمان
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB max


class FileHandler:
    """معالج الملفات - تحميل من Google Drive"""
    
    def load_file(self, file_source):
        """تحميل ملف من Google Drive أو أي مصدر"""
        try:
            # تحويل رابط Google Drive
            if 'drive.google.com' in file_source:
                file_source = self._convert_gdrive_url(file_source)
            
            # تحميل الملف
            response = requests.get(file_source, timeout=30)
            response.raise_for_status()
            file_content = BytesIO(response.content)
            
            # قراءة حسب النوع
            if '.csv' in file_source.lower():
                df = pd.read_csv(file_content, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file_content)
            
            return df
            
        except Exception as e:
            print(f"خطأ في تحميل الملف: {str(e)}")
            return None
    
    def _convert_gdrive_url(self, url):
        """تحويل رابط Google Drive للتنزيل المباشر"""
        file_id = None
        
        # النمط 1: /file/d/FILE_ID/
        match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
        if match:
            file_id = match.group(1)
        
        # النمط 2: id=FILE_ID
        if not file_id:
            match = re.search(r'id=([a-zA-Z0-9_-]+)', url)
            if match:
                file_id = match.group(1)
        
        if file_id:
            return f"https://drive.google.com/uc?export=download&id={file_id}"
        
        return url


class DescriptiveAnalyzer:
    """محرك التحليل الوصفي"""
    
    def __init__(self, dataframe):
        self.df = dataframe
    
    def run_analysis(self):
        """تحليل وصفي كامل لجميع المتغيرات"""
        results = {
            "متغيرات_رقمية": [],
            "متغيرات_فئوية": []
        }
        
        for column in self.df.columns:
            try:
                if pd.api.types.is_numeric_dtype(self.df[column]):
                    # متغير رقمي
                    data = self.df[column].dropna()
                    if len(data) > 0:
                        results["متغيرات_رقمية"].append({
                            "المتغير": column,
                            "العدد": int(len(data)),
                            "المتوسط": round(float(data.mean()), 2),
                            "الوسيط": round(float(data.median()), 2),
                            "الانحراف_المعياري": round(float(data.std()), 2),
                            "أصغر_قيمة": round(float(data.min()), 2),
                            "أكبر_قيمة": round(float(data.max()), 2)
                        })
                else:
                    # متغير فئوي
                    data = self.df[column].dropna()
                    if len(data) > 0:
                        counts = data.value_counts()
                        percentages = (counts / len(data) * 100).round(1)
                        
                        categories = []
                        for cat in counts.index[:5]:  # أول 5 فئات
                            categories.append({
                                "الفئة": str(cat),
                                "التكرار": int(counts[cat]),
                                "النسبة": float(percentages[cat])
                            })
                        
                        results["متغيرات_فئوية"].append({
                            "المتغير": column,
                            "عدد_الفئات": int(len(counts)),
                            "التوزيع": categories
                        })
            except:
                continue
        
        return results


class InferentialAnalyzer:
    """محرك الاختبارات الاستدلالية"""
    
    def __init__(self, dataframe):
        self.df = dataframe
    
    def ttest(self, group_var, value_var):
        """اختبار T للعينات المستقلة"""
        try:
            clean_df = self.df[[group_var, value_var]].dropna()
            groups = clean_df[group_var].unique()
            
            if len(groups) != 2:
                return {"error": f"يجب أن يحتوي {group_var} على فئتين فقط. الفئات الحالية: {len(groups)}"}
            
            group1 = clean_df[clean_df[group_var] == groups[0]][value_var]
            group2 = clean_df[clean_df[group_var] == groups[1]][value_var]
            
            t_stat, p_value = stats.ttest_ind(group1, group2)
            
            # حجم الأثر (Cohen's d)
            pooled_std = np.sqrt(((len(group1)-1)*group1.std()**2 + (len(group2)-1)*group2.std()**2) / (len(group1)+len(group2)-2))
            cohens_d = (group1.mean() - group2.mean()) / pooled_std if pooled_std != 0 else 0
            
            return {
                "المجموعة_1": {
                    "الاسم": str(groups[0]),
                    "العدد": int(len(group1)),
                    "المتوسط": round(float(group1.mean()), 2),
                    "الانحراف": round(float(group1.std()), 2)
                },
                "المجموعة_2": {
                    "الاسم": str(groups[1]),
                    "العدد": int(len(group2)),
                    "المتوسط": round(float(group2.mean()), 2),
                    "الانحراف": round(float(group2.std()), 2)
                },
                "t": round(float(t_stat), 3),
                "p": round(float(p_value), 4),
                "cohens_d": round(float(cohens_d), 3),
                "دال": p_value < 0.05,
                "تفسير": self._interpret_ttest(p_value, cohens_d)
            }
        except Exception as e:
            return {"error": f"خطأ في اختبار T: {str(e)}"}
    
    def _interpret_ttest(self, p, d):
        """تفسير نتائج اختبار T"""
        result = []
        
        if p < 0.001:
            result.append("يوجد فرق دال إحصائياً عند مستوى 0.001 ⭐⭐⭐")
        elif p < 0.01:
            result.append("يوجد فرق دال إحصائياً عند مستوى 0.01 ⭐⭐")
        elif p < 0.05:
            result.append("يوجد فرق دال إحصائياً عند مستوى 0.05 ⭐")
        else:
            result.append("لا يوجد فرق دال إحصائياً ❌")
        
        abs_d = abs(d)
        if abs_d < 0.2:
            result.append("حجم الأثر: ضعيف جداً")
        elif abs_d < 0.5:
            result.append("حجم الأثر: صغير")
        elif abs_d < 0.8:
            result.append("حجم الأثر: متوسط")
        else:
            result.append("حجم الأثر: كبير")
        
        return " | ".join(result)
    
    def anova(self, dependent, independent):
        """تحليل التباين الأحادي (One-Way ANOVA)"""
        try:
            clean_df = self.df[[dependent, independent]].dropna()
            groups = [group[dependent].values for name, group in clean_df.groupby(independent)]
            group_names = clean_df[independent].unique()
            
            f_stat, p_value = stats.f_oneway(*groups)
            
            # إحصائيات المجموعات
            group_stats = []
            for name in group_names:
                g_data = clean_df[clean_df[independent] == name][dependent]
                group_stats.append({
                    "المجموعة": str(name),
                    "العدد": int(len(g_data)),
                    "المتوسط": round(float(g_data.mean()), 2),
                    "الانحراف": round(float(g_data.std()), 2)
                })
            
            return {
                "عدد_المجموعات": len(groups),
                "المجموعات": group_stats,
                "F": round(float(f_stat), 3),
                "p": round(float(p_value), 4),
                "دال": p_value < 0.05,
                "تفسير": "توجد فروق دالة إحصائياً ✅" if p_value < 0.05 else "لا توجد فروق دالة ❌"
            }
        except Exception as e:
            return {"error": f"خطأ في ANOVA: {str(e)}"}
    
    def correlation(self, variables):
        """تحليل الارتباط - بيرسون، سبيرمان، كيندال"""
        try:
            clean_df = self.df[variables].dropna()
            
            results = []
            for i in range(len(variables)):
                for j in range(i+1, len(variables)):
                    var1, var2 = variables[i], variables[j]
                    
                    # بيرسون (Pearson) - الارتباط الخطي
                    r_pearson, p_pearson = stats.pearsonr(clean_df[var1], clean_df[var2])
                    
                    # سبيرمان (Spearman) - الارتباط الرتبي
                    r_spearman, p_spearman = stats.spearmanr(clean_df[var1], clean_df[var2])
                    
                    # كيندال (Kendall) - الارتباط الرتبي
                    r_kendall, p_kendall = stats.kendalltau(clean_df[var1], clean_df[var2])
                    
                    results.append({
                        "المتغير_1": var1,
                        "المتغير_2": var2,
                        "بيرسون_r": round(float(r_pearson), 3),
                        "بيرسون_p": round(float(p_pearson), 4),
                        "بيرسون_دال": "نعم ✅" if p_pearson < 0.05 else "لا ❌",
                        "سبيرمان_rho": round(float(r_spearman), 3),
                        "سبيرمان_p": round(float(p_spearman), 4),
                        "سبيرمان_دال": "نعم ✅" if p_spearman < 0.05 else "لا ❌",
                        "كيندال_tau": round(float(r_kendall), 3),
                        "كيندال_p": round(float(p_kendall), 4),
                        "كيندال_دال": "نعم ✅" if p_kendall < 0.05 else "لا ❌",
                        "القوة": self._interpret_r(r_pearson),
                        "الاتجاه": "طردي ↗️" if r_pearson > 0 else "عكسي ↘️"
                    })
            
            return {"الارتباطات": results}
        except Exception as e:
            return {"error": f"خطأ في الارتباط: {str(e)}"}
    
    def _interpret_r(self, r):
        """تفسير قوة الارتباط"""
        abs_r = abs(r)
        if abs_r < 0.3:
            return "ضعيف"
        elif abs_r < 0.5:
            return "متوسط"
        elif abs_r < 0.7:
            return "قوي"
        else:
            return "قوي جداً"


class RegressionAnalyzer:
    """محرك تحليل الانحدار"""
    
    def __init__(self, dataframe):
        self.df = dataframe
    
    def multiple_regression(self, dependent, independents):
        """تحليل الانحدار المتعدد"""
        try:
            all_vars = [dependent] + independents
            clean_df = self.df[all_vars].dropna()
            
            X = clean_df[independents]
            y = clean_df[dependent]
            X = sm.add_constant(X)
            
            model = sm.OLS(y, X).fit()
            
            # المعاملات
            coefficients = []
            for i, var in enumerate(['الثابت'] + independents):
                coefficients.append({
                    "المتغير": var,
                    "المعامل": round(float(model.params[i]), 4),
                    "t": round(float(model.tvalues[i]), 3),
                    "p": round(float(model.pvalues[i]), 4),
                    "دال": model.pvalues[i] < 0.05
                })
            
            return {
                "R2": round(float(model.rsquared), 4),
                "R2_معدل": round(float(model.rsquared_adj), 4),
                "F": round(float(model.fvalue), 3),
                "p_F": round(float(model.f_pvalue), 4),
                "المعاملات": coefficients,
                "تفسير": f"النموذج يفسر {round(model.rsquared*100, 1)}% من التباين"
            }
        except Exception as e:
            return {"error": f"خطأ في الانحدار: {str(e)}"}


class ReportGenerator:
    """مولد التقارير بالعربية"""
    
    def generate(self, results, analysis_type):
        """توليد تقرير نصي منسق"""
        
        if analysis_type == 'descriptive':
            return self._format_descriptive(results)
        elif analysis_type == 'ttest':
            return self._format_ttest(results)
        elif analysis_type == 'anova':
            return self._format_anova(results)
        elif analysis_type == 'correlation':
            return self._format_correlation(results)
        elif analysis_type == 'regression':
            return self._format_regression(results)
        else:
            return str(results)
    
    def _format_descriptive(self, r):
        """تقرير التحليل الوصفي"""
        report = "📊 *نتائج التحليل الوصفي*\n"
        report += "━━━━━━━━━━━━━━━━━━━━\n\n"
        
        if r.get('متغيرات_رقمية'):
            report += "*🔢 المتغيرات الرقمية:*\n\n"
            for v in r['متغيرات_رقمية']:
                report += f"▫️ *{v['المتغير']}*\n"
                report += f"   • العدد: {v['العدد']}\n"
                report += f"   • المتوسط: {v['المتوسط']}\n"
                report += f"   • الانحراف المعياري: {v['الانحراف_المعياري']}\n"
                report += f"   • المدى: {v['أصغر_قيمة']} - {v['أكبر_قيمة']}\n\n"
        
        if r.get('متغيرات_فئوية'):
            report += "*📝 المتغيرات الفئوية:*\n\n"
            for v in r['متغيرات_فئوية']:
                report += f"▫️ *{v['المتغير']}* ({v['عدد_الفئات']} فئة)\n"
                for cat in v['التوزيع'][:3]:
                    report += f"   • {cat['الفئة']}: {cat['التكرار']} ({cat['النسبة']}%)\n"
                report += "\n"
        
        return report
    
    def _format_ttest(self, r):
        """تقرير اختبار T"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "📊 *نتائج اختبار T*\n"
        report += "━━━━━━━━━━━━━━━━━━━━\n\n"
        
        report += f"*{r['المجموعة_1']['الاسم']}:*\n"
        report += f"• ن = {r['المجموعة_1']['العدد']}\n"
        report += f"• م = {r['المجموعة_1']['المتوسط']}\n"
        report += f"• ع = {r['المجموعة_1']['الانحراف']}\n\n"
        
        report += f"*{r['المجموعة_2']['الاسم']}:*\n"
        report += f"• ن = {r['المجموعة_2']['العدد']}\n"
        report += f"• م = {r['المجموعة_2']['المتوسط']}\n"
        report += f"• ع = {r['المجموعة_2']['الانحراف']}\n\n"
        
        report += "*النتيجة:*\n"
        report += f"• t = {r['t']}\n"
        report += f"• p = {r['p']}\n"
        report += f"• Cohen's d = {r['cohens_d']}\n\n"
        
        report += f"*✅ التفسير:*\n{r['تفسير']}"
        
        return report
    
    def _format_anova(self, r):
        """تقرير ANOVA"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "📊 *نتائج تحليل التباين (ANOVA)*\n"
        report += "━━━━━━━━━━━━━━━━━━━━\n\n"
        
        report += "*المجموعات:*\n"
        for g in r['المجموعات']:
            report += f"• {g['المجموعة']}: ن={g['العدد']}, م={g['المتوسط']}\n"
        
        report += f"\n*النتيجة:*\n"
        report += f"• F = {r['F']}\n"
        report += f"• p = {r['p']}\n\n"
        
        report += f"*✅ التفسير:* {r['تفسير']}"
        
        return report
    
    def _format_correlation(self, r):
        """تقرير الارتباط - بيرسون، سبيرمان، كيندال"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "📊 *نتائج تحليل الارتباط*\n"
        report += "━━━━━━━━━━━━━━━━━━━━\n\n"
        
        for c in r['الارتباطات']:
            report += f"🔗 *{c['المتغير_1']} ↔️ {c['المتغير_2']}*\n"
            report += f"   {c['الاتجاه']} - {c['القوة']}\n\n"
            
            report += f"   📌 *بيرسون (Pearson):*\n"
            report += f"      • r = {c['بيرسون_r']}\n"
            report += f"      • p = {c['بيرسون_p']}\n"
            report += f"      • دال إحصائياً: {c['بيرسون_دال']}\n\n"
            
            report += f"   📌 *سبيرمان (Spearman):*\n"
            report += f"      • rho = {c['سبيرمان_rho']}\n"
            report += f"      • p = {c['سبيرمان_p']}\n"
            report += f"      • دال إحصائياً: {c['سبيرمان_دال']}\n\n"
            
            report += f"   📌 *كيندال (Kendall):*\n"
            report += f"      • tau = {c['كيندال_tau']}\n"
            report += f"      • p = {c['كيندال_p']}\n"
            report += f"      • دال إحصائياً: {c['كيندال_دال']}\n"
            
            report += "\n" + "─"*30 + "\n\n"
        
        return report
    
    def _format_regression(self, r):
        """تقرير الانحدار"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "📊 *نتائج تحليل الانحدار*\n"
        report += "━━━━━━━━━━━━━━━━━━━━\n\n"
        
        report += f"*جودة النموذج:*\n"
        report += f"• R² = {r['R2']} ({round(r['R2']*100, 1)}%)\n"
        report += f"• F = {r['F']} (p = {r['p_F']})\n\n"
        
        report += "*المعاملات:*\n"
        for c in r['المعاملات']:
            sig = "✅" if c['دال'] else "❌"
            report += f"{sig} {c['المتغير']}: β={c['المعامل']} (p={c['p']})\n"
        
        report += f"\n*✅ {r['تفسير']}*"
        
        return report


# ============= API ENDPOINTS =============

@app.route('/')
def home():
    """الصفحة الرئيسية"""
    return jsonify({
        "service": "نظام التحليل الإحصائي الآلي - الجزائر",
        "version": "1.0",
        "status": "active",
        "endpoints": {
            "/health": "GET - فحص الصحة",
            "/analyze": "POST - تحليل البيانات"
        }
    })


@app.route('/health')
def health():
    """فحص صحة الخادم"""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat()
    }), 200


@app.route('/analyze', methods=['POST'])
def analyze():
    """نقطة الدخول الرئيسية للتحليل"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({"success": False, "error": "لم يتم إرسال بيانات"}), 400
        
        # التحقق من المدخلات الأساسية
        if 'file_url' not in data or 'analysis_type' not in data:
            return jsonify({"success": False, "error": "file_url و analysis_type مطلوبان"}), 400
        
        # تحميل الملف
        file_handler = FileHandler()
        df = file_handler.load_file(data['file_url'])
        
        if df is None:
            return jsonify({"success": False, "error": "فشل تحميل الملف. تحقق من الرابط والصلاحيات"}), 400
        
        # تنفيذ التحليل المطلوب
        analysis_type = data['analysis_type'].lower()
        result = None
        
        if analysis_type == 'descriptive':
            analyzer = DescriptiveAnalyzer(df)
            result = analyzer.run_analysis()
        
        elif analysis_type == 'ttest':
            vars = data.get('variables', {})
            analyzer = InferentialAnalyzer(df)
            result = analyzer.ttest(vars.get('group_var'), vars.get('value_var'))
        
        elif analysis_type == 'anova':
            vars = data.get('variables', {})
            analyzer = InferentialAnalyzer(df)
            result = analyzer.anova(vars.get('dependent'), vars.get('independent'))
        
        elif analysis_type == 'correlation':
            vars = data.get('variables', {})
            analyzer = InferentialAnalyzer(df)
            result = analyzer.correlation(vars.get('list', []))
        
        elif analysis_type == 'regression':
            vars = data.get('variables', {})
            analyzer = RegressionAnalyzer(df)
            result = analyzer.multiple_regression(vars.get('dependent'), vars.get('independents', []))
        
        else:
            return jsonify({"success": False, "error": f"نوع التحليل '{analysis_type}' غير مدعوم"}), 400
        
        # توليد التقرير
        report_gen = ReportGenerator()
        report = report_gen.generate(result, analysis_type)
        
        return jsonify({
            "success": True,
            "analysis_type": analysis_type,
            "timestamp": datetime.now().isoformat(),
            "data": result,
            "report": report
        }), 200
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
