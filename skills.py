import openpyxl

class SkillsGenerator:
    def __init__(self, excel_file, output_tex):
        self.excel_file = excel_file
        self.output_tex = output_tex
        self.data = []
        self.load_data()

    def load_data(self):
        wb = openpyxl.load_workbook(self.excel_file)
        ws = wb['skills']  # Make sure the sheet is named 'skills'
        for row in ws.iter_rows(min_row=2, values_only=True):
            skill = {
                "category": row[0],
                "skills": row[1]
            }
            self.data.append(skill)

    def generate_tex(self):
        with open(self.output_tex, 'w') as f:
            f.write(r"""
            %-------------------------------------------------------------------------------
            % CONTENT
            %-------------------------------------------------------------------------------
            \begin{cvskills}
            """)
            
            for skill in self.data:
                f.write(r"""
                %---------------------------------------------------------
                \cvskill
                {""" + skill["category"] + r"""} % Category
                {""" + skill["skills"] + r"""} % Skills
                """)
            
            f.write(r"""
            %---------------------------------------------------------
            \end{cvskills}
            """)

if __name__ == "__main__":
    excel_file = "data.xlsx"  # Replace with your Excel file path
    output_tex = "resume/skills.tex"
    
    generator = SkillsGenerator(excel_file, output_tex)
    generator.generate_tex()
