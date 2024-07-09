import os
import subprocess
import pandas as pd
from flask import Flask, request, render_template, send_file

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/students'
app.config['RESULT_FOLDER'] = 'uploads/results'
app.config['EXAM_FILE'] = '20240706cuoti.xlsx'
app.config['IMAGE_FOLDER'] = 'static/images'

# Load exam data once
exam_data = {}
def load_exam_data():
    global exam_data
    exam_file = app.config['EXAM_FILE']
    xls = pd.ExcelFile(exam_file)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        exam_data[sheet_name] = df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file and file.filename.endswith('.xlsx'):
        filename = file.filename
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        print(f"File saved to: {file_path}")  # Debugging line
        pdf_file_path = process_file(filename)
        return send_file(pdf_file_path, as_attachment=True)
    return redirect(request.url)

def process_file(filename):
    student_file = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    output_pdf_folder = app.config['RESULT_FOLDER']
    student_name = os.path.splitext(filename)[0]

    # Load student errors
    df = pd.read_excel(student_file, dtype={'试卷名称': str, '错误题号': str})
    df['试卷名称'] = df['试卷名称'].apply(lambda x: str(x).rstrip('.0'))
    df['错误题号'] = df['错误题号'].apply(lambda x: str(x).rstrip('.0'))

    # Generate LaTeX content
    latex_content = generate_latex_content(student_name, df, exam_data)
    
    # Create LaTeX file
    latex_template = r"""
    \documentclass{ctexart}[10pt,a4paper]
    \usepackage{geometry}
    \usepackage{fancyhdr}
    \usepackage{pifont}
    \usepackage{graphicx}
    \usepackage{amsmath}
    \usepackage{amssymb}
    \usepackage{amsthm}
    \usepackage{fourier}
    \newtheorem{remark}{注}
    \newtheorem{definition}{定义}
    \theoremstyle{definition}
    \newtheorem{example}{例}
    \newtheorem{zhuyi}{.}
    \newtheorem{theorem}{定理}
    \newcommand\backcong{\mathrel{\reflectbox{$\cong$}}}
    \usepackage{ifthen}
    \newlength{\la}
    \newlength{\lb}
    \newlength{\lc}
    \newlength{\ld}
    \newlength{\lhalf}
    \newlength{\lquarter}
    \newlength{\lmax}
    \newcommand{\xz}{\nolinebreak\dotfill\mbox{\raisebox{-1.8pt}
            {$\cdots$}(\hspace{1cm})}}
    \newcommand{\xx}[4]{\\[.5pt]%
        \settowidth{\la}{A.~#1~~~}
        \settowidth{\lb}{B.~#2~~~}
        \settowidth{\lc}{C.~#3~~~}
        \settowidth{\ld}{D.~#4~~~}
        \ifthenelse{\lengthtest{\la > \lb}}  {\setlength{\lmax}{\la}}  {\setlength{\lmax}{\lb}}
        \ifthenelse{\lengthtest{\lmax < \lc}}  {\setlength{\lmax}{\lc}}  {}
        \ifthenelse{\lengthtest{\lmax < \ld}}  {\setlength{\lmax}{\ld}}  {}
        \setlength{\lhalf}{0.5\linewidth}
        \setlength{\lquarter}{0.25\linewidth}
        \ifthenelse{\lengthtest{\lmax > \lhalf}}  {\noindent{}A.~#1 \\ B.~#2 \\ C.~#3 \\ D.~#4 }  {
            \ifthenelse{\lengthtest{\lmax > \lquarter}}  {\noindent\makebox[\lhalf][l]{A.~#1~~~}%
                \makebox[\lhalf][l]{B.~#2~~~}%
                \\
                \makebox[\lhalf][l]{C.~#3~~~}%
                \makebox[\lhalf][l]{D.~#4~~~}}%
            {\noindent\makebox[\lquarter][l]{A.~#1~~~}%
                \makebox[\lquarter][l]{B.~#2~~~}%
                \makebox[\lquarter][l]{C.~#3~~~}%
                \makebox[\lquarter][l]{D.~#4~~~}}}}
    \newcommand{\tk}[1][2.5]{\,\underline{\mbox{\hspace{#1 cm}}}\,}
    \begin{document}
    \begin{enumerate}
    % INSERT CONTENT HERE
    \end{enumerate}
    \end{document}
    """
    output_tex_file = os.path.join(output_pdf_folder, f"{student_name}.tex").replace('\\', '/')
    output_pdf_file = os.path.join(output_pdf_folder, f"{student_name}.pdf").replace('\\', '/')
    print(f"Generating PDF: {output_pdf_file}")  # Debugging line
    create_latex_file(latex_template, latex_content, output_tex_file)
    compile_latex_to_pdf(output_tex_file, output_pdf_file)

    # Clean up auxiliary files
    aux_extensions = ['.aux', '.log', '.out', '.toc']
    for ext in aux_extensions:
        aux_file = output_tex_file.replace('.tex', ext)
        if os.path.exists(aux_file):
            os.remove(aux_file)
    if os.path.exists(output_tex_file):
        os.remove(output_tex_file)

    # 确保PDF文件已生成
    if os.path.exists(output_pdf_file):
        return output_pdf_file
    else:
        raise FileNotFoundError(f"PDF file not generated: {output_pdf_file}")

def generate_latex_content(student_name, errors_df, exam_data):
    latex_content = ""
    for _, row in errors_df.iterrows():
        exam_name = str(row['试卷名称'])
        wrong_question_numbers = str(row['错误题号']).split('、') if pd.notna(row['错误题号']) else []
        if exam_name in exam_data:
            exam_df = exam_data[exam_name]
            for wrong_question_number in wrong_question_numbers:
                if wrong_question_number.isdigit():
                    wrong_question_number = int(wrong_question_number)
                    if 1 <= wrong_question_number <= len(exam_df):
                        question = exam_df.iloc[wrong_question_number - 1]['题目']
                        question_name = f"[{exam_name}-{wrong_question_number}]"
                        latex_content += f"\\begin{{minipage}}{{\\linewidth}}\n"
                        latex_content += f"{question}\n"
                        latex_content += f"{question_name}\n"
                        image_path = exam_df.iloc[wrong_question_number - 1]['图片']
                        if pd.notna(image_path):
                            padded_exam_name = exam_name
                            padded_question_number = str(wrong_question_number)
                            filled_content = f"{padded_exam_name}-{padded_question_number}"
                            image_file = os.path.join(app.config['IMAGE_FOLDER'], f"{filled_content}.png").replace('\\', '/')
                            if os.path.exists(image_file):
                                latex_content += f"\\vspace{{0.5cm}}\n"
                                latex_content += f"\n"
                                latex_content += f"\\includegraphics[height=3cm]{{{image_file}}}\n"
                            else:
                                latex_content += f"Image {filled_content} does not exist.\n"
                        latex_content += f"\\end{{minipage}}\n\\vspace{{0.5cm}}\n\\noindent\n"
                    else:
                        latex_content += "\n"
                else:
                    latex_content += "\n"
        else:
            latex_content += "\n"
    return latex_content

def create_latex_file(template, content, output_file):
    latex_content = template.replace('% INSERT CONTENT HERE', content)
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(latex_content)

def compile_latex_to_pdf(latex_file, output_pdf_file):
    try:
        result = subprocess.run(
            ['xelatex', '-interaction=nonstopmode', '-output-directory', os.path.dirname(output_pdf_file), latex_file],
            check=True, capture_output=True, text=True
        )
        print(f"LaTeX output: {result.stdout}")
        print(f"LaTeX errors: {result.stderr}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error compiling {latex_file}: {e.stderr}")
        return False

if __name__ == '__main__':
    load_exam_data()
    app.run(host='127.0.0.1', port=5000, debug=True)

