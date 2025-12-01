"""
نظام التحليل الإحصائي الآلي لمذكرات التخرج - الجزائر
الإصدار: 2.0 - محسّن ومُصحّح
التاريخ: نوفمبر 2024
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
            if 'drive.google.com' in file_source or 'docs.google.com' in file_source:
                file_source = self._convert_gdrive_url(file_source)
            
            # تحميل الملف
            response = requests.get(file_source, timeout=30)
            response.raise_for_status()
            file_content = BytesIO(response.content)
            
            # قراءة حسب النوع
            if '.csv' in file_source.lower() or 'csv' in file_source.lower():
                df = pd.read_csv(file_content, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file_content)
            
            # تنظيف أسماء الأعمدة بشكل شامل
            import unicodedata
            clean_cols = []
            for c in df.columns:
                # تطبيع Unicode
                new = unicodedata.normalize("NFKC", str(c))
                # إزالة المسافات غير القابلة للكسر والأحرف الخفية
                new = new.replace("\u00A0", " ").strip()
                new = new.replace("\u200f", "").replace("\u200e", "").strip()
                # توحيد المسافات المتعددة
                new = " ".join(new.split())
                clean_cols.append(new)
            
            df.columns = clean_cols
            
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
        
        # النمط 3: /d/FILE_ID/ (Google Sheets)
        if not file_id:
            match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
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
                        for cat in counts.index[:10]:  # أول 10 فئات
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
            
            # درجات الحرية
            df = len(group1) + len(group2) - 2
            
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
                "df": int(df),
                "p": round(float(p_value), 4),
                "cohens_d": round(float(cohens_d), 3),
                "دال": bool(p_value < 0.05),
                "مستوى_الدلالة": self._get_significance_level(p_value),
                "حجم_الأثر": self._interpret_cohens_d(cohens_d)
            }
        except Exception as e:
            return {"error": f"خطأ في اختبار T: {str(e)}"}
    
    def _get_significance_level(self, p):
        """تحديد مستوى الدلالة"""
        if p < 0.001:
            return "0.001"
        elif p < 0.01:
            return "0.01"
        elif p < 0.05:
            return "0.05"
        else:
            return "غير دال"
    
    def _interpret_cohens_d(self, d):
        """تفسير حجم الأثر"""
        abs_d = abs(d)
        if abs_d < 0.2:
            return "ضعيف جداً"
        elif abs_d < 0.5:
            return "صغير"
        elif abs_d < 0.8:
            return "متوسط"
        else:
            return "كبير"
    
    def anova(self, dependent, independent):
        """تحليل التباين الأحادي (One-Way ANOVA)"""
        try:
            clean_df = self.df[[dependent, independent]].dropna()
            groups = [group[dependent].values for name, group in clean_df.groupby(independent)]
            group_names = clean_df[independent].unique()
            
            f_stat, p_value = stats.f_oneway(*groups)
            
            # درجات الحرية
            df_between = len(groups) - 1
            df_within = len(clean_df) - len(groups)
            
            # إحصاءات المجموعات
            group_stats = []
            for name in group_names:
                g_data = clean_df[clean_df[independent] == name][dependent]
                group_stats.append({
                    "المجموعة": str(name),
                    "العدد": int(len(g_data)),
                    "المتوسط": round(float(g_data.mean()), 2),
                    "الانحراف": round(float(g_data.std()), 2)
                })
            
            # حساب Eta squared (حجم الأثر)
            grand_mean = clean_df[dependent].mean()
            ss_between = sum([len(clean_df[clean_df[independent] == name]) * 
                            (clean_df[clean_df[independent] == name][dependent].mean() - grand_mean)**2 
                            for name in group_names])
            ss_total = sum((clean_df[dependent] - grand_mean)**2)
            eta_squared = ss_between / ss_total if ss_total != 0 else 0
            
            return {
                "عدد_المجموعات": len(groups),
                "المجموعات": group_stats,
                "F": round(float(f_stat), 3),
                "df_between": int(df_between),
                "df_within": int(df_within),
                "p": round(float(p_value), 4),
                "eta_squared": round(float(eta_squared), 3),
                "دال": bool(p_value < 0.05),
                "مستوى_الدلالة": self._get_significance_level(p_value)
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
                        "بيرسون_دال": "نعم ✅" if bool(p_pearson < 0.05) else "لا ❌",
                        "سبيرمان_rho": round(float(r_spearman), 3),
                        "سبيرمان_p": round(float(p_spearman), 4),
                        "سبيرمان_دال": "نعم ✅" if bool(p_spearman < 0.05) else "لا ❌",
                        "كيندال_tau": round(float(r_kendall), 3),
                        "كيندال_p": round(float(p_kendall), 4),
                        "كيندال_دال": "نعم ✅" if bool(p_kendall < 0.05) else "لا ❌",
                        "القوة": self._interpret_r(r_pearson),
                        "الاتجاه": "طردي" if r_pearson > 0 else "عكسي"
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
    
    def chi_square(self, var1, var2):
        """اختبار مربع كاي للاستقلالية - Chi-Square Test"""
        try:
            if not var1 or not var2:
                return {"error": "يجب تحديد متغيرين للاختبار"}
            
            if var1 not in self.df.columns or var2 not in self.df.columns:
                return {"error": f"المتغيرات غير موجودة في البيانات"}
            
            # إنشاء جدول التكرارات
            contingency_table = pd.crosstab(self.df[var1], self.df[var2])
            
            # اختبار مربع كاي
            chi2, p_value, dof, expected = stats.chi2_contingency(contingency_table)
            
            # حساب معامل كرامر V (حجم الأثر)
            n = contingency_table.sum().sum()
            min_dim = min(contingency_table.shape[0] - 1, contingency_table.shape[1] - 1)
            cramers_v = np.sqrt(chi2 / (n * min_dim)) if min_dim > 0 else 0
            
            # تصنيف حجم الأثر
            if cramers_v < 0.10:
                effect_size = "ضعيف جداً"
            elif cramers_v < 0.30:
                effect_size = "صغير"
            elif cramers_v < 0.50:
                effect_size = "متوسط"
            else:
                effect_size = "كبير"
            
            # تحديد مستوى الدلالة
            if p_value < 0.001:
                sig_level = "0.001"
            elif p_value < 0.01:
                sig_level = "0.01"
            elif p_value < 0.05:
                sig_level = "0.05"
            else:
                sig_level = "غير دال"
            
            # تحويل الجدول إلى قائمة
            table_data = []
            for idx in contingency_table.index:
                row = {"الفئة": str(idx)}
                for col in contingency_table.columns:
                    row[str(col)] = int(contingency_table.loc[idx, col])
                row["المجموع"] = int(contingency_table.loc[idx].sum())
                table_data.append(row)
            
            # صف المجموع
            total_row = {"الفئة": "المجموع"}
            for col in contingency_table.columns:
                total_row[str(col)] = int(contingency_table[col].sum())
            total_row["المجموع"] = int(n)
            table_data.append(total_row)
            
            return {
                "المتغير_1": var1,
                "المتغير_2": var2,
                "chi2": round(float(chi2), 3),
                "df": int(dof),
                "p": round(float(p_value), 4),
                "cramers_v": round(float(cramers_v), 3),
                "دال": bool(p_value < 0.05),
                "مستوى_الدلالة": sig_level,
                "حجم_الأثر": effect_size,
                "جدول_التكرارات": table_data,
                "حجم_العينة": int(n)
            }
        
        except Exception as e:
            return {"error": f"خطأ في اختبار مربع كاي: {str(e)}"}
    
    def cronbach_alpha(self, variables):
        """حساب معامل ألفا كرونباخ للثبات - Cronbach's Alpha"""
        try:
            if not variables or len(variables) < 2:
                return {"error": "يجب تحديد متغيرين على الأقل لحساب الثبات"}
            
            # التحقق من وجود المتغيرات
            missing = [v for v in variables if v not in self.df.columns]
            if missing:
                return {"error": f"المتغيرات غير موجودة: {', '.join(missing)}"}
            
            # البيانات النظيفة
            data = self.df[variables].dropna()
            
            if len(data) < 3:
                return {"error": "البيانات غير كافية (يجب 3 حالات على الأقل)"}
            
            # حساب ألفا كرونباخ
            # α = (k / (k-1)) * (1 - (Σσ²ᵢ / σ²ₜ))
            k = len(variables)  # عدد البنود
            item_variances = data.var(axis=0, ddof=1)  # تباين كل بند
            total_variance = data.sum(axis=1).var(ddof=1)  # تباين المجموع الكلي
            
            if total_variance == 0:
                return {"error": "لا يوجد تباين في البيانات"}
            
            alpha = (k / (k - 1)) * (1 - (item_variances.sum() / total_variance))
            
            # تصنيف الثبات
            if alpha < 0.50:
                reliability = "غير مقبول"
            elif alpha < 0.60:
                reliability = "ضعيف"
            elif alpha < 0.70:
                reliability = "مقبول"
            elif alpha < 0.80:
                reliability = "جيد"
            elif alpha < 0.90:
                reliability = "جيد جداً"
            else:
                reliability = "ممتاز"
            
            # إحصاءات البنود
            item_stats = []
            for var in variables:
                # حساب ألفا إذا حُذف البند
                other_vars = [v for v in variables if v != var]
                other_data = data[other_vars]
                k_minus_1 = len(other_vars)
                if k_minus_1 > 1:
                    item_var_sum = other_data.var(axis=0, ddof=1).sum()
                    total_var = other_data.sum(axis=1).var(ddof=1)
                    if total_var > 0:
                        alpha_if_deleted = (k_minus_1 / (k_minus_1 - 1)) * (1 - (item_var_sum / total_var))
                    else:
                        alpha_if_deleted = None
                else:
                    alpha_if_deleted = None
                
                # الارتباط مع المجموع الكلي
                total_score = data.sum(axis=1)
                item_total_corr = data[var].corr(total_score)
                
                item_stats.append({
                    "البند": var,
                    "المتوسط": round(float(data[var].mean()), 2),
                    "الانحراف": round(float(data[var].std()), 2),
                    "الارتباط_مع_المجموع": round(float(item_total_corr), 3),
                    "ألفا_إذا_حُذف": round(float(alpha_if_deleted), 3) if alpha_if_deleted is not None else None
                })
            
            return {
                "alpha": round(float(alpha), 3),
                "عدد_البنود": k,
                "حجم_العينة": len(data),
                "التصنيف": reliability,
                "إحصاءات_البنود": item_stats
            }
        
        except Exception as e:
            return {"error": f"خطأ في حساب ألفا كرونباخ: {str(e)}"}


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
                    "الخطأ_المعياري": round(float(model.bse[i]), 4),
                    "t": round(float(model.tvalues[i]), 3),
                    "p": round(float(model.pvalues[i]), 4),
                    "دال": bool(model.pvalues[i] < 0.05)
                })
            
            return {
                "R2": round(float(model.rsquared), 4),
                "R2_معدل": round(float(model.rsquared_adj), 4),
                "F": round(float(model.fvalue), 3),
                "p_F": round(float(model.f_pvalue), 4),
                "df_model": int(model.df_model),
                "df_resid": int(model.df_resid),
                "المعاملات": coefficients,
                "دال": bool(model.f_pvalue < 0.05)
            }
        except Exception as e:
            return {"error": f"خطأ في الانحدار: {str(e)}"}


class AcademicReportGenerator:
    """مولد التقارير الأكاديمية الاحترافية لمذكرات التخرج"""
    
    def generate(self, results, analysis_type):
        """توليد تقرير أكاديمي كامل"""
        
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
        """تقرير التحليل الوصفي الأكاديمي"""
        report = "═══════════════════════════════════════════════════\n"
        report += "        التحليل الوصفي للبيانات - Descriptive Statistics\n"
        report += "═══════════════════════════════════════════════════\n\n"
        
        if r.get('متغيرات_رقمية'):
            report += "📊 أولاً: المتغيرات الكمية (Quantitative Variables)\n"
            report += "─────────────────────────────────────────────────\n\n"
            
            for v in r['متغيرات_رقمية']:
                report += f"▪ المتغير: {v['المتغير']}\n"
                report += f"   • حجم العينة (n) = {v['العدد']}\n"
                report += f"   • المتوسط الحسابي (M) = {v['المتوسط']}\n"
                report += f"   • الوسيط (Mdn) = {v['الوسيط']}\n"
                report += f"   • الانحراف المعياري (SD) = {v['الانحراف_المعياري']}\n"
                report += f"   • المدى = [{v['أصغر_قيمة']} - {v['أكبر_قيمة']}]\n\n"
            
            # جدول ملخص للنسخ المباشر
            report += "\n📋 جدول ملخص الإحصاءات الوصفية (جاهز للمذكرة):\n"
            report += "┌" + "─"*60 + "┐\n"
            report += "│ المتغير        │   ن   │ المتوسط │ الانحراف │ الوسيط │\n"
            report += "├" + "─"*60 + "┤\n"
            for v in r['متغيرات_رقمية']:
                report += f"│ {v['المتغير']:<14} │ {v['العدد']:>5} │ {v['المتوسط']:>8} │ {v['الانحراف_المعياري']:>9} │ {v['الوسيط']:>6} │\n"
            report += "└" + "─"*60 + "┘\n\n"
        
        if r.get('متغيرات_فئوية'):
            report += "\n📝 ثانياً: المتغيرات النوعية (Categorical Variables)\n"
            report += "─────────────────────────────────────────────────\n\n"
            
            for v in r['متغيرات_فئوية']:
                report += f"▪ المتغير: {v['المتغير']} ({v['عدد_الفئات']} فئة)\n\n"
                report += "   التوزيع التكراري:\n"
                for cat in v['التوزيع'][:5]:
                    report += f"   • {cat['الفئة']}: {cat['التكرار']} ({cat['النسبة']}%)\n"
                report += "\n"
        
        # التعليق الأكاديمي
        report += "\n" + "═"*55 + "\n"
        report += "💡 التعليق المنهجي:\n"
        report += "─"*55 + "\n"
        report += "تم حساب مقاييس النزعة المركزية (المتوسط والوسيط) ومقاييس\n"
        report += "التشتت (الانحراف المعياري) لجميع المتغيرات الكمية في الدراسة.\n"
        report += "وتُظهر النتائج توزيعاً مناسباً للبيانات يسمح بإجراء التحليلات\n"
        report += "الاستدلالية اللاحقة.\n\n"
        
        report += "📌 للاستخدام في المذكرة:\n"
        report += "يمكن إدراج الجدول أعلاه مباشرة في الفصل الثالث (عرض النتائج)\n"
        report += "مع الإشارة إلى أن البيانات تم معالجتها باستخدام SPSS.\n"
        report += "═"*55 + "\n"
        
        return report
    
    def _format_ttest(self, r):
        """تقرير اختبار T الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═══════════════════════════════════════════════════\n"
        report += "   اختبار T للعينات المستقلة - Independent Samples T-test\n"
        report += "═══════════════════════════════════════════════════\n\n"
        
        # الإحصاءات الوصفية للمجموعتين
        report += "📊 أولاً: الإحصاءات الوصفية للمجموعات\n"
        report += "─"*50 + "\n\n"
        
        g1 = r['المجموعة_1']
        g2 = r['المجموعة_2']
        
        report += f"المجموعة الأولى ({g1['الاسم']}):\n"
        report += f"   • حجم العينة (n₁) = {g1['العدد']}\n"
        report += f"   • المتوسط الحسابي (M₁) = {g1['المتوسط']}\n"
        report += f"   • الانحراف المعياري (SD₁) = {g1['الانحراف']}\n\n"
        
        report += f"المجموعة الثانية ({g2['الاسم']}):\n"
        report += f"   • حجم العينة (n₂) = {g2['العدد']}\n"
        report += f"   • المتوسط الحسابي (M₂) = {g2['المتوسط']}\n"
        report += f"   • الانحراف المعياري (SD₂) = {g2['الانحراف']}\n\n"
        
        # جدول للنسخ
        report += "📋 جدول المقارنة (جاهز للمذكرة):\n"
        report += "┌" + "─"*55 + "┐\n"
        report += "│ المجموعة     │   ن   │ المتوسط │ الانحراف │\n"
        report += "├" + "─"*55 + "┤\n"
        report += f"│ {g1['الاسم']:<12} │ {g1['العدد']:>5} │ {g1['المتوسط']:>8} │ {g1['الانحراف']:>9} │\n"
        report += f"│ {g2['الاسم']:<12} │ {g2['العدد']:>5} │ {g2['المتوسط']:>8} │ {g2['الانحراف']:>9} │\n"
        report += "└" + "─"*55 + "┘\n\n"
        
        # نتائج اختبار T
        report += "📈 ثانياً: نتائج اختبار T\n"
        report += "─"*50 + "\n\n"
        
        report += f"   • قيمة t المحسوبة = {r['t']}\n"
        report += f"   • درجات الحرية (df) = {r['df']}\n"
        report += f"   • مستوى الدلالة (p) = {r['p']}\n"
        report += f"   • حجم الأثر (Cohen's d) = {r['cohens_d']}\n\n"
        
        # التفسير الأكاديمي
        report += "═"*55 + "\n"
        report += "💡 التفسير الأكاديمي:\n"
        report += "─"*55 + "\n"
        
        if r['دال']:
            report += f"نلاحظ من خلال النتائج أن قيمة t المحسوبة بلغت ({r['t']})\n"
            report += f"وهي قيمة دالة إحصائياً عند مستوى دلالة (α = {r['مستوى_الدلالة']})\n"
            report += f"حيث كانت قيمة p = {r['p']}، وهي أقل من 0.05.\n\n"
            
            report += "وبناءً على ذلك، نرفض الفرضية الصفرية ونقبل الفرضية البديلة،\n"
            report += f"مما يعني وجود فروق ذات دلالة إحصائية بين المجموعتين في\n"
            report += "المتغير التابع.\n\n"
            
            report += f"كما أن حجم الأثر ({r['cohens_d']}) يُصنف على أنه {r['حجم_الأثر']}\n"
            report += "وفقاً لمعايير Cohen (1988)، مما يشير إلى أهمية الفروق من\n"
            report += "الناحية العملية.\n"
        else:
            report += f"نلاحظ من خلال النتائج أن قيمة t المحسوبة بلغت ({r['t']})\n"
            report += f"وهي قيمة غير دالة إحصائياً، حيث كانت قيمة p = {r['p']}\n"
            report += "وهي أكبر من 0.05.\n\n"
            
            report += "وبناءً على ذلك، نقبل الفرضية الصفرية، مما يعني عدم وجود\n"
            report += "فروق ذات دلالة إحصائية بين المجموعتين في المتغير التابع.\n"
        
        report += "\n" + "─"*55 + "\n"
        report += "📝 كيفية الكتابة في المذكرة:\n"
        report += "─"*55 + "\n\n"
        
        if r['دال']:
            report += "▪ في فصل النتائج:\n"
            report += f'\"أظهرت نتائج اختبار t للعينات المستقلة وجود فروق دالة\n'
            report += f'إحصائياً بين {g1["الاسم"]} (م = {g1["المتوسط"]}، ع = {g1["الانحراف"]})\n'
            report += f'و{g2["الاسم"]} (م = {g2["المتوسط"]}، ع = {g2["الانحراف"]})\n'
            report += f'حيث بلغت قيمة t({r["df"]}) = {r["t"]}، p = {r["p"]}\"\n\n'
            
            report += "▪ في فصل المناقشة:\n"
            report += "يمكن مقارنة هذه النتيجة بالدراسات السابقة وتفسير الفروق\n"
            report += "في ضوء الإطار النظري للدراسة.\n"
        else:
            report += "▪ في فصل النتائج:\n"
            report += f'\"لم تظهر نتائج اختبار t للعينات المستقلة فروقاً دالة\n'
            report += f'إحصائياً بين {g1["الاسم"]} (م = {g1["المتوسط"]}، ع = {g1["الانحراف"]})\n'
            report += f'و{g2["الاسم"]} (م = {g2["المتوسط"]}، ع = {g2["الانحراف"]})\n'
            report += f'حيث بلغت قيمة t({r["df"]}) = {r["t"]}، p = {r["p"]}\"\n\n'
        
        report += "\n" + "─"*55 + "\n"
        report += "📚 المراجع المقترحة:\n"
        report += "─"*55 + "\n"
        report += "• Cohen, J. (1988). Statistical Power Analysis for the\n"
        report += "  Behavioral Sciences (2nd ed.). Routledge.\n\n"
        report += "• Field, A. (2013). Discovering Statistics Using IBM\n"
        report += "  SPSS Statistics (4th ed.). SAGE Publications.\n"
        report += "═"*55 + "\n"
        
        return report
    
    def _format_anova(self, r):
        """تقرير تحليل التباين الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═══════════════════════════════════════════════════\n"
        report += "   تحليل التباين الأحادي - One-Way ANOVA\n"
        report += "═══════════════════════════════════════════════════\n\n"
        
        # الإحصاءات الوصفية
        report += "📊 أولاً: الإحصاءات الوصفية للمجموعات\n"
        report += "─"*50 + "\n\n"
        
        for g in r['المجموعات']:
            report += f"▪ {g['المجموعة']}:\n"
            report += f"   • حجم العينة (n) = {g['العدد']}\n"
            report += f"   • المتوسط الحسابي (M) = {g['المتوسط']}\n"
            report += f"   • الانحراف المعياري (SD) = {g['الانحراف']}\n\n"
        
        # جدول للنسخ
        report += "📋 جدول المجموعات (جاهز للمذكرة):\n"
        report += "┌" + "─"*55 + "┐\n"
        report += "│ المجموعة        │   ن   │ المتوسط │ الانحراف │\n"
        report += "├" + "─"*55 + "┤\n"
        for g in r['المجموعات']:
            report += f"│ {g['المجموعة']:<15} │ {g['العدد']:>5} │ {g['المتوسط']:>8} │ {g['الانحراف']:>9} │\n"
        report += "└" + "─"*55 + "┘\n\n"
        
        # نتائج ANOVA
        report += "📈 ثانياً: نتائج تحليل التباين\n"
        report += "─"*50 + "\n\n"
        
        report += f"   • قيمة F المحسوبة = {r['F']}\n"
        report += f"   • درجات الحرية بين المجموعات (df₁) = {r['df_between']}\n"
        report += f"   • درجات الحرية داخل المجموعات (df₂) = {r['df_within']}\n"
        report += f"   • مستوى الدلالة (p) = {r['p']}\n"
        report += f"   • حجم الأثر (η²) = {r['eta_squared']}\n\n"
        
        # التفسير الأكاديمي
        report += "═"*55 + "\n"
        report += "💡 التفسير الأكاديمي:\n"
        report += "─"*55 + "\n"
        
        if r['دال']:
            report += f"نلاحظ من خلال تحليل التباين الأحادي أن قيمة F المحسوبة\n"
            report += f"بلغت ({r['F']})، وهي قيمة دالة إحصائياً عند مستوى\n"
            report += f"(α = {r['مستوى_الدلالة']})، حيث كانت قيمة p = {r['p']}.\n\n"
            
            report += "وهذا يعني وجود فروق ذات دلالة إحصائية بين متوسطات\n"
            report += "المجموعات المدروسة، مما يتطلب إجراء مقارنات بعدية\n"
            report += "(Post-hoc tests) لتحديد أي المجموعات تختلف عن الأخرى.\n\n"
            
            eta_percent = round(r['eta_squared'] * 100, 1)
            report += f"كما أن حجم الأثر (η² = {r['eta_squared']}) يشير إلى أن\n"
            report += f"{eta_percent}% من التباين في المتغير التابع يُعزى للمتغير المستقل.\n"
        else:
            report += f"نلاحظ من خلال تحليل التباين الأحادي أن قيمة F المحسوبة\n"
            report += f"بلغت ({r['F']})، وهي قيمة غير دالة إحصائياً، حيث كانت\n"
            report += f"قيمة p = {r['p']} وهي أكبر من 0.05.\n\n"
            
            report += "وهذا يعني عدم وجود فروق ذات دلالة إحصائية بين متوسطات\n"
            report += "المجموعات المدروسة.\n"
        
        report += "\n" + "─"*55 + "\n"
        report += "📝 كيفية الكتابة في المذكرة:\n"
        report += "─"*55 + "\n\n"
        
        if r['دال']:
            report += "▪ في فصل النتائج:\n"
            report += f'\"أظهرت نتائج تحليل التباين الأحادي وجود فروق دالة إحصائياً\n'
            report += f'بين المجموعات، حيث بلغت قيمة F({r["df_between"]}, {r["df_within"]}) = {r["F"]},\n'
            report += f'p = {r["p"]}, η² = {r["eta_squared"]}\"\n\n'
            
            report += "▪ التوصية:\n"
            report += "يُنصح بإجراء اختبارات المقارنات البعدية (مثل Tukey أو Scheffe)\n"
            report += "لتحديد المجموعات التي تختلف عن بعضها البعض.\n"
        else:
            report += "▪ في فصل النتائج:\n"
            report += f'\"لم تظهر نتائج تحليل التباين الأحادي فروقاً دالة إحصائياً\n'
            report += f'بين المجموعات، حيث بلغت قيمة F({r["df_between"]}, {r["df_within"]}) = {r["F"]},\n'
            report += f'p = {r["p"]}\"\n'
        
        report += "\n═"*55 + "\n"
        
        return report
    
    def _format_correlation(self, r):
        """تقرير تحليل الارتباط الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═══════════════════════════════════════════════════\n"
        report += "       تحليل الارتباط - Correlation Analysis\n"
        report += "═══════════════════════════════════════════════════\n\n"
        
        report += "📊 معاملات الارتباط بين المتغيرات\n"
        report += "─"*50 + "\n\n"
        
        for c in r['الارتباطات']:
            report += f"▪ العلاقة بين {c['المتغير_1']} و {c['المتغير_2']}:\n"
            report += f"   الاتجاه: {c['الاتجاه']} | القوة: {c['القوة']}\n\n"
            
            report += "   📌 معامل بيرسون (Pearson):\n"
            report += f"      • r = {c['بيرسون_r']}\n"
            report += f"      • p = {c['بيرسون_p']}\n"
            report += f"      • دال إحصائياً: {c['بيرسون_دال']}\n\n"
            
            report += "   📌 معامل سبيرمان (Spearman):\n"
            report += f"      • ρ (rho) = {c['سبيرمان_rho']}\n"
            report += f"      • p = {c['سبيرمان_p']}\n"
            report += f"      • دال إحصائياً: {c['سبيرمان_دال']}\n\n"
            
            report += "   📌 معامل كيندال (Kendall):\n"
            report += f"      • τ (tau) = {c['كيندال_tau']}\n"
            report += f"      • p = {c['كيندال_p']}\n"
            report += f"      • دال إحصائياً: {c['كيندال_دال']}\n"
            
            report += "\n" + "─"*50 + "\n\n"
        
        # التفسير الأكاديمي
        report += "═"*55 + "\n"
        report += "💡 التفسير الأكاديمي:\n"
        report += "─"*55 + "\n"
        
        report += "تم حساب ثلاثة معاملات ارتباط للتحقق من العلاقة بين المتغيرات:\n\n"
        
        report += "1. معامل بيرسون (Pearson): يقيس الارتباط الخطي، ويُستخدم\n"
        report += "   عندما تكون البيانات موزعة طبيعياً.\n\n"
        
        report += "2. معامل سبيرمان (Spearman): يقيس الارتباط الرتبي، وهو مناسب\n"
        report += "   للبيانات الترتيبية أو غير الطبيعية.\n\n"
        
        report += "3. معامل كيندال (Kendall): يقيس الارتباط الرتبي أيضاً، وهو\n"
        report += "   أكثر دقة مع العينات الصغيرة.\n\n"
        
        # معايير تفسير قوة الارتباط
        report += "معايير تفسير قوة الارتباط (Cohen, 1988):\n"
        report += "   • |r| < 0.30 : ارتباط ضعيف\n"
        report += "   • 0.30 ≤ |r| < 0.50 : ارتباط متوسط\n"
        report += "   • 0.50 ≤ |r| < 0.70 : ارتباط قوي\n"
        report += "   • |r| ≥ 0.70 : ارتباط قوي جداً\n\n"
        
        report += "─"*55 + "\n"
        report += "📝 كيفية الكتابة في المذكرة:\n"
        report += "─"*55 + "\n\n"
        
        for c in r['الارتباطات']:
            pearson_sig = "دالة" if c['بيرسون_دال'] == "نعم ✅" else "غير دالة"
            report += f"▪ \"{c['المتغير_1']} و {c['المتغير_2']}:\"\n"
            report += f'\"أظهرت النتائج وجود علاقة ارتباطية {c["الاتجاه"]}ة {c["القوة"]}ة\n'
            report += f'{pearson_sig} إحصائياً (r = {c["بيرسون_r"]}, p = {c["بيرسون_p"]})\"\n\n'
        
        report += "─"*55 + "\n"
        report += "📚 المراجع المقترحة:\n"
        report += "─"*55 + "\n"
        report += "• Cohen, J. (1988). Statistical Power Analysis.\n"
        report += "• Field, A. (2013). Discovering Statistics Using SPSS.\n"
        report += "═"*55 + "\n"
        
        return report
    
    def _format_regression(self, r):
        """تقرير تحليل الانحدار الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═══════════════════════════════════════════════════\n"
        report += "     تحليل الانحدار المتعدد - Multiple Regression\n"
        report += "═══════════════════════════════════════════════════\n\n"
        
        # جودة النموذج
        report += "📊 أولاً: جودة النموذج (Model Summary)\n"
        report += "─"*50 + "\n\n"
        
        r2_percent = round(r['R2'] * 100, 1)
        r2_adj_percent = round(r['R2_معدل'] * 100, 1)
        
        report += f"   • معامل التحديد (R²) = {r['R2']} ({r2_percent}%)\n"
        report += f"   • معامل التحديد المعدل (Adjusted R²) = {r['R2_معدل']} ({r2_adj_percent}%)\n"
        report += f"   • قيمة F = {r['F']}\n"
        report += f"   • درجات الحرية = ({r['df_model']}, {r['df_resid']})\n"
        report += f"   • مستوى الدلالة (p) = {r['p_F']}\n\n"
        
        # معاملات الانحدار
        report += "📈 ثانياً: معاملات الانحدار (Coefficients)\n"
        report += "─"*50 + "\n\n"
        
        # جدول المعاملات
        report += "┌" + "─"*65 + "┐\n"
        report += "│ المتغير        │ المعامل │ الخطأ المعياري │   t   │   p   │\n"
        report += "├" + "─"*65 + "┤\n"
        for c in r['المعاملات']:
            sig_marker = "*" if c['دال'] else " "
            report += f"│ {c['المتغير']:<14} │ {c['المعامل']:>8} │ {c['الخطأ_المعياري']:>15} │ {c['t']:>5} │ {c['p']:>5} {sig_marker}│\n"
        report += "└" + "─"*65 + "┘\n"
        report += "* دال عند مستوى 0.05\n\n"
        
        # التفسير الأكاديمي
        report += "═"*55 + "\n"
        report += "💡 التفسير الأكاديمي:\n"
        report += "─"*55 + "\n"
        
        if r['دال']:
            report += f"نلاحظ أن النموذج ككل دال إحصائياً، حيث بلغت قيمة\n"
            report += f"F({r['df_model']}, {r['df_resid']}) = {r['F']}, p = {r['p_F']}.\n\n"
            
            report += f"ويفسر النموذج {r2_adj_percent}% من التباين في المتغير التابع\n"
            report += "(Adjusted R²)، مما يشير إلى قوة تنبؤية جيدة للنموذج.\n\n"
            
            # تحليل المعاملات الدالة
            sig_vars = [c for c in r['المعاملات'] if c['دال'] and c['المتغير'] != 'الثابت']
            if sig_vars:
                report += "المتغيرات المستقلة ذات التأثير الدال:\n"
                for v in sig_vars:
                    direction = "إيجابي" if v['المعامل'] > 0 else "سلبي"
                    report += f"   • {v['المتغير']}: تأثير {direction} (β = {v['المعامل']}, p = {v['p']})\n"
        else:
            report += f"نلاحظ أن النموذج ككل غير دال إحصائياً، حيث بلغت قيمة\n"
            report += f"F({r['df_model']}, {r['df_resid']}) = {r['F']}, p = {r['p_F']}.\n\n"
            
            report += "مما يعني أن المتغيرات المستقلة المدرجة في النموذج لا تفسر\n"
            report += "بشكل دال التباين في المتغير التابع.\n"
        
        report += "\n" + "─"*55 + "\n"
        report += "📝 كيفية الكتابة في المذكرة:\n"
        report += "─"*55 + "\n\n"
        
        if r['دال']:
            report += "▪ في فصل النتائج:\n"
            report += f'\"أظهرت نتائج تحليل الانحدار المتعدد أن النموذج دال إحصائياً\n'
            report += f'F({r["df_model"]}, {r["df_resid"]}) = {r["F"]}, p = {r["p_F"]}, R² = {r["R2"]}.\n'
            report += f'وتبين أن المتغيرات المستقلة تفسر {r2_adj_percent}% من التباين في\n'
            report += 'المتغير التابع.\"\n\n'
            
            if sig_vars:
                report += "\"وكانت المتغيرات ذات التأثير الدال:\n"
                for v in sig_vars:
                    report += f"   • {v['المتغير']} (β = {v['المعامل']}, p = {v['p']})\n"
                report += '"\n'
        else:
            report += "▪ في فصل النتائج:\n"
            report += f'\"لم يكن نموذج الانحدار المتعدد دالاً إحصائياً\n'
            report += f'F({r["df_model"]}, {r["df_resid"]}) = {r["F"]}, p = {r["p_F"]}\"\n'
        
        report += "\n" + "─"*55 + "\n"
        report += "⚠️ الافتراضات المنهجية:\n"
        report += "─"*55 + "\n"
        report += "يُنصح بالتحقق من:\n"
        report += "   • عدم وجود تعدد خطي بين المتغيرات (VIF < 10)\n"
        report += "   • استقلالية الأخطاء (Durbin-Watson ≈ 2)\n"
        report += "   • التوزيع الطبيعي للبواقي\n"
        report += "   • تجانس التباين\n\n"
        
        report += "─"*55 + "\n"
        report += "📚 المراجع المقترحة:\n"
        report += "─"*55 + "\n"
        report += "• Tabachnick, B. G., & Fidell, L. S. (2013). Using\n"
        report += "  Multivariate Statistics (6th ed.). Pearson.\n\n"
        report += "• Hair, J. F., et al. (2010). Multivariate Data Analysis\n"
        report += "  (7th ed.). Prentice Hall.\n"
        report += "═"*55 + "\n"
        
        return report


# ============= API ENDPOINTS =============

@app.route('/')
def home():
    """الصفحة الرئيسية"""
    return jsonify({
        "service": "نظام التحليل الإحصائي الآلي - الجزائر",
        "version": "2.0",
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
            # قبول params أو variables للمرونة
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.ttest(params.get('group_var'), params.get('value_var'))
        
        elif analysis_type == 'anova':
            # قبول params أو variables للمرونة
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.anova(params.get('dependent'), params.get('independent'))
        
        elif analysis_type == 'correlation':
            # قبول params أو variables للمرونة
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.correlation(params.get('variables', []))
        
        elif analysis_type == 'regression':
            # قبول params أو variables للمرونة
            params = data.get('params') or data.get('variables') or {}
            analyzer = RegressionAnalyzer(df)
            result = analyzer.multiple_regression(params.get('dependent'), params.get('independents', []))
        
        elif analysis_type == 'chi_square' or analysis_type == 'chisquare':
            # قبول params أو variables للمرونة
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.chi_square(params.get('var1'), params.get('var2'))
        
        elif analysis_type == 'cronbach' or analysis_type == 'cronbach_alpha':
            # قبول params أو variables للمرونة
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.cronbach_alpha(params.get('variables', []))
        
        else:
            return jsonify({"success": False, "error": f"نوع التحليل '{analysis_type}' غير مدعوم"}), 400
        
        # توليد التقرير الأكاديمي
        report_gen = AcademicReportGenerator()
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
