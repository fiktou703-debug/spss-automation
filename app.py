"""
نظام التحليل الإحصائي الآلي لمذكرات التخرج - الجزائر
الإصدار: 2.4 - مع دعم Word Generator
التاريخ: ديسمبر 2024

التعديلات الجديدة:
- إضافة endpoint جديد /analyze_word لتوليد ملفات Word
- دعم كامل لجميع أنواع التحليلات السبعة
- تنسيق احترافي حسب المعايير الأكاديمية الجزائرية
"""

from flask import Flask, request, jsonify, send_file
import pandas as pd
import numpy as np
from scipy import stats
import statsmodels.api as sm
from datetime import datetime
import requests
from io import BytesIO
import re
import traceback
import tempfile
import os

# Import Word Generator
from spss_word_generator import SPSSWordGenerator

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
            return "ضعيف"
        elif abs_d < 0.8:
            return "متوسط"
        else:
            return "كبير"
    
    def anova(self, dependent, independent):
        """تحليل التباين الأحادي"""
        try:
            clean_df = self.df[[independent, dependent]].dropna()
            groups = []
            labels = []
            
            for name, group in clean_df.groupby(independent):
                groups.append(group[dependent].values)
                labels.append(name)
            
            if len(groups) < 2:
                return {"error": f"يجب وجود مجموعتين على الأقل في {independent}"}
            
            # تحليل التباين
            f_stat, p_value = stats.f_oneway(*groups)
            
            # حساب مجموع المربعات
            grand_mean = clean_df[dependent].mean()
            ss_between = sum([len(g) * (np.mean(g) - grand_mean)**2 for g in groups])
            ss_within = sum([np.sum((g - np.mean(g))**2) for g in groups])
            ss_total = ss_between + ss_within
            
            # درجات الحرية
            df_between = len(groups) - 1
            df_within = len(clean_df) - len(groups)
            df_total = len(clean_df) - 1
            
            # متوسط المربعات
            ms_between = ss_between / df_between
            ms_within = ss_within / df_within
            
            # حجم الأثر (Eta Squared)
            eta_squared = ss_between / ss_total
            
            # ===== NEW: إحصاءات المجموعات =====
            group_descriptives = {}
            for i, name in enumerate(labels):
                group_data = groups[i]
                group_descriptives[str(name)] = {
                    'العدد': int(len(group_data)),
                    'المتوسط': round(float(np.mean(group_data)), 2),
                    'الانحراف_المعياري': round(float(np.std(group_data, ddof=1)), 2)
                }
            
            return {
                "N": int(len(clean_df)),
                "إحصاءات_المجموعات": group_descriptives,
                "بين_المجموعات": {
                    "مجموع_المربعات": round(float(ss_between), 3),
                    "درجات_الحرية": int(df_between),
                    "متوسط_المربعات": round(float(ms_between), 3)
                },
                "داخل_المجموعات": {
                    "مجموع_المربعات": round(float(ss_within), 3),
                    "درجات_الحرية": int(df_within),
                    "متوسط_المربعات": round(float(ms_within), 3)
                },
                "الكلي": {
                    "مجموع_المربعات": round(float(ss_total), 3),
                    "درجات_الحرية": int(df_total)
                },
                "F": round(float(f_stat), 3),
                "p": round(float(p_value), 4),
                "eta_squared": round(float(eta_squared), 3),
                "دال": bool(p_value < 0.05),
                "مستوى_الدلالة": self._get_significance_level(p_value),
                "حجم_الأثر": self._interpret_eta_squared(eta_squared)
            }
        except Exception as e:
            return {"error": f"خطأ في ANOVA: {str(e)}"}

    def _interpret_eta_squared(self, eta):
        """تفسير Eta Squared"""
        if eta < 0.01:
            return "ضعيف جداً"
        elif eta < 0.06:
            return "ضعيف"
        elif eta < 0.14:
            return "متوسط"
        else:
            return "كبير"
    
    def correlation(self, variables):
        """تحليل الارتباط"""
        try:
            if not variables or len(variables) < 2:
                return {"error": "يجب تحديد متغيرين على الأقل"}
            
            # استخراج البيانات
            data = self.df[variables].dropna()
            
            if len(data) < 3:
                return {"error": "عدد المشاهدات غير كافٍ (أقل من 3)"}
            
            # ===== NEW: إحصاءات وصفية =====
            descriptive_stats = {}
            for var in variables:
                descriptive_stats[var] = {
                    'N': int(len(data)),
                    'Mean': round(float(data[var].mean()), 2),
                    'SD': round(float(data[var].std(ddof=1)), 2)
                }
            
            # حساب الارتباط
            corr_matrix = data.corr()
            
            # بناء المصفوفة مع قيم p - FIXED KEYS
            result_matrix = {}
            significant_results = []
            
            for var1 in variables:
                result_matrix[var1] = {}
                for var2 in variables:
                    if var1 == var2:
                        result_matrix[var1][var2] = {
                            "r": 1.0,
                            "p": 0.0
                        }
                    else:
                        r, p = stats.pearsonr(data[var1], data[var2])
                        result_matrix[var1][var2] = {
                            "r": round(float(r), 3),
                            "p": round(float(p), 4)
                        }
                        
                        # جمع النتائج الدالة
                        if p < 0.05 and var1 < var2:  # تجنب التكرار
                            significant_results.append({
                                'var1': var1,
                                'var2': var2,
                                'r': round(float(r), 3),
                                'p': round(float(p), 4),
                                'قوة': self._interpret_correlation_strength(abs(r))
                            })
            
            return {
                "method": "pearson",
                "N": int(len(data)),
                "إحصاءات_وصفية": descriptive_stats,
                "مصفوفة_الارتباط": result_matrix,
                "نتائج_دالة": significant_results
            }
        except Exception as e:
            return {"error": f"خطأ في تحليل الارتباط: {str(e)}"}
    
    def _interpret_correlation_strength(self, abs_r):
        """تفسير قوة الارتباط"""
        if abs_r < 0.3:
            return "ضعيفة"
        elif abs_r < 0.5:
            return "متوسطة"
        elif abs_r < 0.7:
            return "قوية"
        else:
            return "قوية جداً"

    def chi_square(self, var1, var2):
        """اختبار مربع كاي"""
        try:
            # Create contingency table
            contingency = pd.crosstab(self.df[var1], self.df[var2])
            
            # Chi-square test
            chi2, p, dof, expected = stats.chi2_contingency(contingency)
            
            # Cramér's V
            n = contingency.sum().sum()
            min_dim = min(contingency.shape[0], contingency.shape[1]) - 1
            cramers_v = np.sqrt(chi2 / (n * min_dim))
            
            return {
                "N": int(n),
                "var1": var1,
                "var2": var2,
                "chi_square": round(float(chi2), 3),
                "df": int(dof),
                "p": round(float(p), 4),
                "cramers_v": round(float(cramers_v), 3),
                "دال": bool(p < 0.05),
                "مستوى_الدلالة": self._get_significance_level(p),
                "قوة_العلاقة": self._interpret_cramers_v(cramers_v),
                "جدول_التوافق": contingency.to_dict()
            }
        except Exception as e:
            return {"error": f"خطأ في Chi-Square: {str(e)}"}

    def _interpret_cramers_v(self, v):
        """تفسير Cramér's V"""
        if v < 0.1:
            return "ضعيف جداً"
        elif v < 0.3:
            return "ضعيف"
        elif v < 0.5:
            return "متوسط"
        else:
            return "قوي"
    
    def cronbach_alpha(self, variables):
        """حساب معامل ألفا كرونباخ"""
        try:
            if not variables or len(variables) < 2:
                return {"error": "يجب تحديد متغيرين على الأقل"}
            
            # استخراج البيانات
            data = self.df[variables].dropna()
            
            if len(data) < 2:
                return {"error": "عدد المشاهدات غير كافٍ"}
            
            # حساب Cronbach's Alpha
            item_vars = data.var(axis=0, ddof=1)
            total_var = data.sum(axis=1).var(ddof=1)
            n_items = len(variables)
            
            alpha = (n_items / (n_items - 1)) * (1 - item_vars.sum() / total_var)
            
            # إحصاءات البنود
            items_stats = []
            for var in variables:
                # Alpha if item deleted
                other_vars = [v for v in variables if v != var]
                if len(other_vars) > 1:
                    temp_data = data[other_vars]
                    temp_item_vars = temp_data.var(axis=0, ddof=1)
                    temp_total_var = temp_data.sum(axis=1).var(ddof=1)
                    n_temp = len(other_vars)
                    alpha_if_deleted = (n_temp / (n_temp - 1)) * (1 - temp_item_vars.sum() / temp_total_var)
                else:
                    alpha_if_deleted = None
                
                items_stats.append({
                    "البند": var,
                    "المتوسط": round(float(data[var].mean()), 2),
                    "الانحراف": round(float(data[var].std()), 2),
                    "الارتباط_مع_المجموع": round(float(data[var].corr(data.sum(axis=1))), 3),
                    "ألفا_إذا_حُذف": round(float(alpha_if_deleted), 3) if alpha_if_deleted else None
                })
            
            return {
                "alpha": round(float(alpha), 3),
                "عدد_البنود": n_items,
                "حجم_العينة": int(len(data)),
                "التصنيف": self._classify_alpha(alpha),
                "إحصاءات_البنود": items_stats
            }
        except Exception as e:
            return {"error": f"خطأ في Cronbach's Alpha: {str(e)}"}
    
    def _classify_alpha(self, alpha):
        """تصنيف قيمة Alpha"""
        if alpha >= 0.9:
            return "ممتاز (Excellent)"
        elif alpha >= 0.8:
            return "جيد (Good)"
        elif alpha >= 0.7:
            return "مقبول (Acceptable)"
        elif alpha >= 0.6:
            return "مشكوك فيه (Questionable)"
        elif alpha >= 0.5:
            return "ضعيف (Poor)"
        else:
            return "غير مقبول (Unacceptable)"


class RegressionAnalyzer:
    """محرك تحليل الانحدار"""
    
    def __init__(self, dataframe):
        self.df = dataframe
    
    def multiple_regression(self, dependent, independents):
        """تحليل الانحدار المتعدد"""
        try:
            if not independents or len(independents) < 1:
                return {"error": "يجب تحديد متغير مستقل واحد على الأقل"}
            
            # إعداد البيانات
            cols = [dependent] + independents
            data = self.df[cols].dropna()
            
            if len(data) < len(independents) + 2:
                return {"error": "عدد المشاهدات غير كافٍ للانحدار"}
            
            # إعداد المتغيرات
            X = data[independents]
            y = data[dependent]
            X = sm.add_constant(X)  # إضافة الثابت
            
            # تشغيل الانحدار
            model = sm.OLS(y, X).fit()
            
            # استخراج النتائج
            coefficients = []
            for i, var in enumerate(['Constant'] + independents):
                coefficients.append({
                    "المتغير": var,
                    "المعامل": round(float(model.params[i]), 3),
                    "الخطأ_المعياري": round(float(model.bse[i]), 3),
                    "t": round(float(model.tvalues[i]), 3),
                    "p": round(float(model.pvalues[i]), 4)
                })
            
            return {
                "R": round(float(np.sqrt(model.rsquared)), 3),
                "R2": round(float(model.rsquared), 3),
                "R2_المعدل": round(float(model.rsquared_adj), 3),
                "الخطأ_المعياري": round(float(np.sqrt(model.mse_resid)), 3),
                "F": round(float(model.fvalue), 3),
                "p_model": round(float(model.f_pvalue), 4),
                "دال": bool(model.f_pvalue < 0.05),
                "معاملات": coefficients
            }
        except Exception as e:
            return {"error": f"خطأ في الانحدار: {str(e)}"}


class AcademicReportGenerator:
    """مولد التقارير الأكاديمية النصية (ASCII format)"""
    
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
        elif analysis_type in ['chi_square', 'chisquare']:
            return self._format_chisquare(results)
        elif analysis_type in ['cronbach', 'cronbach_alpha']:
            return self._format_cronbach(results)
        else:
            return "نوع التحليل غير مدعوم"
    
    def _format_descriptive(self, r):
        """تقرير الإحصاء الوصفي الأكاديمي"""
        report = "═"*55 + "\n"
        report += "        التحليل الإحصائي الوصفي\n"
        report += "        Descriptive Statistics Analysis\n"
        report += "═"*55 + "\n\n"
        
        # المتغيرات الرقمية
        if r.get('متغيرات_رقمية'):
            report += "📊 أولاً: الإحصاءات الوصفية للمتغيرات الرقمية\n"
            report += "─"*55 + "\n\n"
            
            report += "┌" + "─"*70 + "┐\n"
            report += "│ المتغير       │  N   │  Mean  │  SD   │  Min  │  Max  │\n"
            report += "├" + "─"*70 + "┤\n"
            
            for var in r['متغيرات_رقمية']:
                report += f"│ {var['المتغير']:<14} │ {var['العدد']:>4} │ {var['المتوسط']:>6.2f} │ {var['الانحراف_المعياري']:>5.2f} │ {var['أصغر_قيمة']:>5.2f} │ {var['أكبر_قيمة']:>5.2f} │\n"
            
            report += "└" + "─"*70 + "┘\n\n"
        
        return report
    
    def _format_ttest(self, r):
        """تقرير اختبار T الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═"*55 + "\n"
        report += "   اختبار T للعينات المستقلة\n"
        report += "   Independent Samples T-Test\n"
        report += "═"*55 + "\n\n"
        
        report += "📊 أولاً: إحصاءات المجموعات\n"
        report += "─"*55 + "\n\n"
        
        report += f"   المجموعة 1: {r['المجموعة_1']['الاسم']}\n"
        report += f"   • العدد (N) = {r['المجموعة_1']['العدد']}\n"
        report += f"   • المتوسط (M) = {r['المجموعة_1']['المتوسط']}\n"
        report += f"   • الانحراف المعياري (SD) = {r['المجموعة_1']['الانحراف']}\n\n"
        
        report += f"   المجموعة 2: {r['المجموعة_2']['الاسم']}\n"
        report += f"   • العدد (N) = {r['المجموعة_2']['العدد']}\n"
        report += f"   • المتوسط (M) = {r['المجموعة_2']['المتوسط']}\n"
        report += f"   • الانحراف المعياري (SD) = {r['المجموعة_2']['الانحراف']}\n\n"
        
        report += "📈 ثانياً: نتائج اختبار T\n"
        report += "─"*55 + "\n\n"
        
        report += f"   • قيمة t = {r['t']}\n"
        report += f"   • درجات الحرية (df) = {r['df']}\n"
        report += f"   • مستوى الدلالة (p) = {r['p']}\n"
        report += f"   • حجم الأثر (Cohen's d) = {r['cohens_d']} ({r['حجم_الأثر']})\n"
        report += f"   • النتيجة: {'دال إحصائياً' if r['دال'] else 'غير دال إحصائياً'}\n\n"
        
        return report
    
    def _format_anova(self, r):
        """تقرير ANOVA الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═"*55 + "\n"
        report += "     تحليل التباين الأحادي - One-Way ANOVA\n"
        report += "═"*55 + "\n\n"
        
        report += "📊 جدول تحليل التباين:\n"
        report += "─"*55 + "\n\n"
        
        report += "┌" + "─"*70 + "┐\n"
        report += "│ مصدر التباين  │    SS    │  df │    MS   │    F   │   Sig. │\n"
        report += "├" + "─"*70 + "┤\n"
        
        report += f"│ بين المجموعات │ {r['بين_المجموعات']['مجموع_المربعات']:>8.3f} │ {r['بين_المجموعات']['درجات_الحرية']:>3} │ {r['بين_المجموعات']['متوسط_المربعات']:>7.3f} │ {r['F']:>6.3f} │ {r['p']:>6.4f} │\n"
        report += f"│ داخل المجموعات│ {r['داخل_المجموعات']['مجموع_المربعات']:>8.3f} │ {r['داخل_المجموعات']['درجات_الحرية']:>3} │ {r['داخل_المجموعات']['متوسط_المربعات']:>7.3f} │    -   │    -   │\n"
        report += f"│ المجموع        │ {r['الكلي']['مجموع_المربعات']:>8.3f} │ {r['الكلي']['درجات_الحرية']:>3} │    -    │    -   │    -   │\n"
        
        report += "└" + "─"*70 + "┘\n\n"
        
        report += f"• حجم الأثر (Eta²) = {r['eta_squared']} ({r['حجم_الأثر']})\n"
        report += f"• النتيجة: {'دال إحصائياً' if r['دال'] else 'غير دال إحصائياً'}\n\n"
        
        return report
    
    def _format_correlation(self, r):
        """تقرير الارتباط الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═"*55 + "\n"
        report += "       تحليل الارتباط - Correlation Analysis\n"
        report += "═"*55 + "\n\n"
        
        report += f"📊 مصفوفة الارتباط (الطريقة: {r['الطريقة'].title()})\n"
        report += f"   عدد المشاهدات: {r['عدد_المشاهدات']}\n"
        report += "─"*55 + "\n\n"
        
        # عرض مبسط للمصفوفة
        report += "(انظر الجداول أعلاه للتفاصيل الكاملة)\n\n"
        
        return report
    
    def _format_regression(self, r):
        """تقرير الانحدار الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═"*55 + "\n"
        report += "  تحليل الانحدار المتعدد\n"
        report += "  Multiple Regression Analysis\n"
        report += "═"*55 + "\n\n"
        
        report += "📊 ملخص النموذج:\n"
        report += "─"*55 + "\n\n"
        
        report += f"   • R = {r['R']}\n"
        report += f"   • R² = {r['R2']}\n"
        report += f"   • R² المعدل = {r['R2_المعدل']}\n"
        report += f"   • الخطأ المعياري = {r['الخطأ_المعياري']}\n\n"
        
        report += "📈 معنوية النموذج:\n"
        report += "─"*55 + "\n\n"
        
        report += f"   • F = {r['F']}\n"
        report += f"   • Sig. = {r['p_model']}\n"
        report += f"   • النتيجة: {'النموذج دال إحصائياً' if r['دال'] else 'النموذج غير دال'}\n\n"
        
        report += "📋 معاملات الانحدار:\n"
        report += "─"*55 + "\n\n"
        
        for coef in r['معاملات']:
            report += f"   {coef['المتغير']}:\n"
            report += f"   • B = {coef['المعامل']}, t = {coef.get('t', 'N/A')}, p = {coef['p']}\n\n"
        
        return report
    
    def _format_chisquare(self, r):
        """تقرير Chi-Square الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═"*55 + "\n"
        report += "   اختبار مربع كاي - Chi-Square Test\n"
        report += "═"*55 + "\n\n"
        
        report += f"📊 نتائج اختبار χ²:\n"
        report += "─"*55 + "\n\n"
        
        report += f"   • χ² = {r['chi2']}\n"
        report += f"   • df = {r['df']}\n"
        report += f"   • Sig. = {r['p']}\n"
        report += f"   • Cramér's V = {r['cramers_v']} ({r['قوة_العلاقة']})\n"
        report += f"   • النتيجة: {'علاقة دالة إحصائياً' if r['دال'] else 'لا توجد علاقة دالة'}\n\n"
        
        return report
    
    def _format_cronbach(self, r):
        """تقرير معامل ألفا كرونباخ الأكاديمي"""
        if 'error' in r:
            return f"❌ خطأ: {r['error']}"
        
        report = "═"*55 + "\n"
        report += "   معامل ألفا كرونباخ للثبات - Cronbach's Alpha\n"
        report += "═"*55 + "\n\n"
        
        report += "📊 أولاً: معامل الثبات العام\n"
        report += "─"*55 + "\n\n"
        
        report += f"   • معامل ألفا (α) = {r['alpha']}\n"
        report += f"   • عدد البنود = {r['عدد_البنود']}\n"
        report += f"   • حجم العينة (N) = {r['حجم_العينة']}\n"
        report += f"   • التصنيف: {r['التصنيف']}\n\n"
        
        report += "📋 ثانياً: جدول إحصاءات البنود\n"
        report += "─"*55 + "\n\n"
        
        report += "┌" + "─"*70 + "┐\n"
        report += "│ البند        │ المتوسط │ الانحراف │ الارتباط │ α إذا حُذف │\n"
        report += "├" + "─"*70 + "┤\n"
        
        for item in r['إحصاءات_البنود']:
            alpha_del = f"{item['ألفا_إذا_حُذف']}" if item['ألفا_إذا_حُذف'] is not None else "N/A"
            report += f"│ {item['البند']:<12} │ {item['المتوسط']:>8} │ {item['الانحراف']:>9} │ {item['الارتباط_مع_المجموع']:>9} │ {alpha_del:>10} │\n"
        
        report += "└" + "─"*70 + "┘\n\n"
        
        return report


# ============= API ENDPOINTS =============

@app.route('/')
def home():
    """الصفحة الرئيسية"""
    return jsonify({
        "service": "نظام التحليل الإحصائي الآلي - الجزائر",
        "version": "2.4",
        "status": "active",
        "endpoints": {
            "/health": "GET - فحص الصحة",
            "/analyze": "POST - تحليل البيانات (JSON + ASCII)",
            "/analyze_word": "POST - تحليل البيانات (Word Document)"
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
    """نقطة الدخول الرئيسية للتحليل - JSON Response"""
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
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.ttest(params.get('group_var'), params.get('value_var'))
        
        elif analysis_type == 'anova':
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.anova(params.get('dependent'), params.get('independent'))
        
        elif analysis_type == 'correlation':
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.correlation(params.get('variables', []))
        
        elif analysis_type == 'regression':
            params = data.get('params') or data.get('variables') or {}
            analyzer = RegressionAnalyzer(df)
            result = analyzer.multiple_regression(params.get('dependent'), params.get('independents', []))
        
        elif analysis_type == 'chi_square' or analysis_type == 'chisquare':
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.chi_square(params.get('var1'), params.get('var2'))
        
        elif analysis_type == 'cronbach' or analysis_type == 'cronbach_alpha':
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


@app.route('/analyze_word', methods=['POST'])
def analyze_word():
    """
    نقطة الدخول الجديدة - Word Document Response
    
    يستقبل نفس البيانات مثل /analyze لكن يرجع ملف Word بدلاً من JSON
    
    Expected JSON input:
    {
        "file_url": "https://...",
        "analysis_type": "descriptive|ttest|anova|correlation|regression|chi_square|cronbach",
        "params" or "variables": {...}
    }
    
    Returns: Word document (.docx)
    """
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
            return jsonify({"success": False, "error": "فشل تحميل الملف"}), 400
        
        # تنفيذ التحليل المطلوب
        analysis_type = data['analysis_type'].lower()
        result = None
        
        if analysis_type == 'descriptive':
            analyzer = DescriptiveAnalyzer(df)
            result = analyzer.run_analysis()
        
        elif analysis_type == 'ttest':
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.ttest(params.get('group_var'), params.get('value_var'))
        
        elif analysis_type == 'anova':
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.anova(params.get('dependent'), params.get('independent'))
        
        elif analysis_type == 'correlation':
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.correlation(params.get('variables', []))
        
        elif analysis_type == 'regression':
            params = data.get('params') or data.get('variables') or {}
            analyzer = RegressionAnalyzer(df)
            result = analyzer.multiple_regression(params.get('dependent'), params.get('independents', []))
        
        elif analysis_type == 'chi_square' or analysis_type == 'chisquare':
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.chi_square(params.get('var1'), params.get('var2'))
        
        elif analysis_type == 'cronbach' or analysis_type == 'cronbach_alpha':
            params = data.get('params') or data.get('variables') or {}
            analyzer = InferentialAnalyzer(df)
            result = analyzer.cronbach_alpha(params.get('variables', []))
        
        else:
            return jsonify({"success": False, "error": f"نوع التحليل '{analysis_type}' غير مدعوم"}), 400
        
        # توليد Word Document
        word_gen = SPSSWordGenerator()
        
        if analysis_type == 'descriptive':
            word_gen.generate_descriptive(result)
        elif analysis_type == 'ttest':
            word_gen.generate_ttest(result)
        elif analysis_type == 'anova':
            word_gen.generate_anova(result)
        elif analysis_type == 'correlation':
            word_gen.generate_correlation(result)
        elif analysis_type == 'regression':
            word_gen.generate_regression(result)
        elif analysis_type in ['chi_square', 'chisquare']:
            word_gen.generate_chisquare(result)
        elif analysis_type in ['cronbach', 'cronbach_alpha']:
            word_gen.generate_cronbach(result)
        
        # حفظ في ملف مؤقت
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        word_gen.save(temp_file.name)
        temp_file.close()
        
        # تحديد اسم الملف
        filename = f"SPSS_{analysis_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        # إرسال الملف
        return send_file(
            temp_file.name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500
    finally:
        # تنظيف الملف المؤقت
        try:
            if 'temp_file' in locals():
                os.unlink(temp_file.name)
        except:
            pass


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
