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
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB max

# -------------------- File Handler --------------------
class FileHandler:
    def load_file(self, file_source):
        try:
            if 'drive.google.com' in file_source:
                file_source = self._convert_gdrive_url(file_source)
            response = requests.get(file_source, timeout=30)
            response.raise_for_status()
            file_content = BytesIO(response.content)
            if '.csv' in file_source.lower():
                df = pd.read_csv(file_content, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file_content)
            return df
        except Exception as e:
            print(f"خطأ في تحميل الملف: {str(e)}")
            return None

    def _convert_gdrive_url(self, url):
        file_id = None
        match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
        if match:
            file_id = match.group(1)
        if not file_id:
            match = re.search(r'id=([a-zA-Z0-9_-]+)', url)
            if match:
                file_id = match.group(1)
        if file_id:
            return f"https://drive.google.com/uc?export=download&id={file_id}"
        return url

# -------------------- Descriptive Analyzer --------------------
class DescriptiveAnalyzer:
    def __init__(self, dataframe):
        self.df = dataframe

    def run_analysis(self):
        results = {"متغيرات_رقمية": [], "متغيرات_فئوية": []}
        for column in self.df.columns:
            try:
                if pd.api.types.is_numeric_dtype(self.df[column]):
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
                    data = self.df[column].dropna()
                    if len(data) > 0:
                        counts = data.value_counts()
                        percentages = (counts / len(data) * 100).round(1)
                        categories = [{"الفئة": str(cat), "التكرار": int(counts[cat]), "النسبة": float(percentages[cat])} for cat in counts.index[:5]]
                        results["متغيرات_فئوية"].append({"المتغير": column, "عدد_الفئات": int(len(counts)), "التوزيع": categories})
            except:
                continue
        return results

# -------------------- Inferential Analyzer --------------------
class InferentialAnalyzer:
    def __init__(self, dataframe):
        self.df = dataframe

    def ttest(self, group_var, value_var):
        try:
            clean_df = self.df[[group_var, value_var]].dropna()
            groups = clean_df[group_var].unique()
            if len(groups) != 2:
                return {"error": f"يجب أن يحتوي {group_var} على فئتين فقط. الفئات الحالية: {len(groups)}"}
            group1 = clean_df[clean_df[group_var] == groups[0]][value_var]
            group2 = clean_df[clean_df[group_var] == groups[1]][value_var]
            t_stat, p_value = stats.ttest_ind(group1, group2)
            pooled_std = np.sqrt(((len(group1)-1)*group1.std()**2 + (len(group2)-1)*group2.std()**2) / (len(group1)+len(group2)-2))
            cohens_d = (group1.mean() - group2.mean()) / pooled_std if pooled_std != 0 else 0
            return {
                "المجموعة_1": {"الاسم": str(groups[0]), "العدد": int(len(group1)), "المتوسط": round(float(group1.mean()), 2), "الانحراف": round(float(group1.std()), 2)},
                "المجموعة_2": {"الاسم": str(groups[1]), "العدد": int(len(group2)), "المتوسط": round(float(group2.mean()), 2), "الانحراف": round(float(group2.std()), 2)},
                "t": round(float(t_stat), 3),
                "p": round(float(p_value), 4),
                "cohens_d": round(float(cohens_d), 3),
                "دال": p_value < 0.05,
                "تفسير": self._interpret_ttest(p_value, cohens_d)
            }
        except Exception as e:
            return {"error": f"خطأ في اختبار T: {str(e)}"}

    def _interpret_ttest(self, p, d):
        result = []
        if p < 0.001: result.append("يوجد فرق دال إحصائياً عند مستوى 0.001 ⭐⭐⭐")
        elif p < 0.01: result.append("يوجد فرق دال إحصائياً عند مستوى 0.01 ⭐⭐")
        elif p < 0.05: result.append("يوجد فرق دال إحصائياً عند مستوى 0.05 ⭐")
        else: result.append("لا يوجد فرق دال إحصائياً ❌")
        abs_d = abs(d)
        if abs_d < 0.2: result.append("حجم الأثر: ضعيف جداً")
        elif abs_d < 0.5: result.append("حجم الأثر: صغير")
        elif abs_d < 0.8: result.append("حجم الأثر: متوسط")
        else: result.append("حجم الأثر: كبير")
        return " | ".join(result)

    def anova(self, dependent, independent):
        try:
            clean_df = self.df[[dependent, independent]].dropna()
            groups = [group[dependent].values for name, group in clean_df.groupby(independent)]
            group_names = clean_df[independent].unique()
            f_stat, p_value = stats.f_oneway(*groups)
            group_stats = [{"المجموعة": str(name), "العدد": int(len(clean_df[clean_df[independent]==name][dependent])), "المتوسط": round(float(clean_df[clean_df[independent]==name][dependent].mean()), 2), "الانحراف": round(float(clean_df[clean_df[independent]==name][dependent].std()),2)} for name in group_names]
            return {"عدد_المجموعات": len(groups), "المجموعات": group_stats, "F": round(float(f_stat),3), "p": round(float(p_value),4), "دال": p_value<0.05, "تفسير": "توجد فروق دالة إحصائياً ✅" if p_value<0.05 else "لا توجد فروق دالة ❌"}
        except Exception as e:
            return {"error": f"خطأ في ANOVA: {str(e)}"}

    def correlation(self, variables):
        try:
            clean_df = self.df[variables].dropna()
            corr_matrix = clean_df.corr()
            results = []
            for i in range(len(variables)):
                for j in range(i+1, len(variables)):
                    var1, var2 = variables[i], variables[j]
                    r = corr_matrix.loc[var1, var2]
                    n = len(clean_df)
                    if abs(r)<1:
                        t = r * np.sqrt((n-2)/(1-r**2))
                        p = 2*(1 - stats.t.cdf(abs(t), n-2))
                    else:
                        p = 0
                    results.append({"المتغير_1": var1, "المتغير_2": var2, "r": round(float(r),3), "p": round(float(p),4), "دال": p<0.05, "القوة": self._interpret_r(r), "الاتجاه": "طردي ↗️" if r>0 else "عكسي ↘️"})
            return {"الارتباطات": results}
        except Exception as e:
            return {"error": f"خطأ في الارتباط: {str(e)}"}

    def _interpret_r(self, r):
        abs_r = abs(r)
        if abs_r<0.3: return "ضعيف"
        elif abs_r<0.5: return "متوسط"
        elif abs_r<0.7: return "قوي"
        else: return "قوي جداً"

# -------------------- Regression Analyzer --------------------
class RegressionAnalyzer:
    def __init__(self, dataframe):
        self.df = dataframe

    def multiple_regression(self, dependent, independents):
        try:
            all_vars = [dependent]+independents
            clean_df = self.df[all_vars].dropna()
            X = sm.add_constant(clean_df[independents])
            y = clean_df[dependent]
            model = sm.OLS(y,X).fit()
            coefficients = [{"المتغير": var, "المعامل": round(float(model.params[i]),4), "t": round(float(model.tvalues[i]),3), "p": round(float(model.pvalues[i]),4), "دال": model.pvalues[i]<0.05} for i,var in enumerate(['الثابت']+independents)]
            return {"R2": round(float(model.rsquared),4), "R2_معدل": round(float(model.rsquared_adj),4), "F": round(float(model.fvalue),3), "p_F": round(float(model.f_pvalue),4), "المعاملات": coefficients, "تفسير": f"النموذج يفسر {round(model.rsquared*100,1)}% من التباين"}
        except Exception as e:
            return {"error": f"خطأ في الانحدار: {str(e)}"}

# -------------------- Report Generator --------------------
class ReportGenerator:
    def generate(self, results, analysis_type):
        if analysis_type=='descriptive': return self._format_descriptive(results)
        elif analysis_type=='ttest': return self._format_ttest(results)
        elif analysis_type=='anova': return self._format_anova(results)
        elif analysis_type=='correlation': return self._format_correlation(results)
        elif analysis_type=='regression': return self._format_regression(results)
        else: return str(results)
    
    # -------------- تنسيقات التقارير (كما في نسخة الأصل) ----------------
    # هنا احتفظ بنفس كل `_format_...` كما أرسلتها أنت
    # لتقليل الحجم، لن أعيد كتابتها بالكامل، فقط انسخها كما هي لديك

# -------------------- API Endpoints --------------------
@app.route('/')
def home():
    return jsonify({"service":"نظام التحليل الإحصائي الآلي - الجزائر","version":"1.0","status":"active","endpoints":{"/health":"GET - فحص الصحة","/analyze":"POST - تحليل البيانات"}})

@app.route('/health')
def health():
    return jsonify({"status":"healthy","timestamp":datetime.now().isoformat()}),200

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        data = request.get_json()
        if not data: return jsonify({"success":False,"error":"لم يتم إرسال بيانات"}),400
        if 'file_url' not in data or 'analysis_type' not in data: return jsonify({"success":False,"error":"file_url و analysis_type مطلوبان"}),400
        df = FileHandler().load_file(data['file_url'])
        if df is None: return jsonify({"success":False,"error":"فشل تحميل الملف. تحقق من الرابط والصلاحيات"}),400

        analysis_type = data['analysis_type'].lower()
        vars = data.get('variables',{})
        if analysis_type=='descriptive': result=DescriptiveAnalyzer(df).run_analysis()
        elif analysis_type=='ttest': result=InferentialAnalyzer(df).ttest(vars.get('group_var'),vars.get('value_var'))
        elif analysis_type=='anova': result=InferentialAnalyzer(df).anova(vars.get('dependent'),vars.get('independent'))
        elif analysis_type=='correlation': result=InferentialAnalyzer(df).correlation(vars.get('list',[]))
        elif analysis_type=='regression': result=RegressionAnalyzer(df).multiple_regression(vars.get('dependent'),vars.get('independents',[]))
        else: return jsonify({"success":False,"error":f"نوع التحليل '{analysis_type}' غير مدعوم"}),400

        report = ReportGenerator().generate(result, analysis_type)
        return jsonify({"success":True,"analysis_type":analysis_type,"timestamp":datetime.now().isoformat(),"data":result,"report":report}),200
    except Exception as e:
        return jsonify({"success":False,"error":str(e),"traceback":traceback.format_exc()}),500

# -------------------- Main --------------------
if __name__=='__main__':
    import os
    port=int(os.environ.get('PORT',5000))
    app.run(host='0.0.0.0',port=port,debug=False)
