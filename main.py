import os
from profile import ResumeGenerator
from summary import SummaryGenerator
from competencies import CompetenciesGenerator
from experience import ExperienceGenerator
from skills import SkillsGenerator
from education import EducationGenerator

def main(excel_file):
    # Extract the base name of the Excel file (without the extension)
    base_name = os.path.splitext(os.path.basename(excel_file))[0]

    # Create the main folder with the name of the Excel file
    main_folder = base_name
    os.makedirs(main_folder, exist_ok=True)

    # Create the 'resume' subfolder within the main folder
    resume_folder = os.path.join(main_folder, 'resume')
    os.makedirs(resume_folder, exist_ok=True)

    # Generate 'main_resume.tex' in the main folder
    main_resume_path = os.path.join(main_folder, 'main_resume.tex')
    with open(main_resume_path, 'w') as f:
        f.write(r"""
        %-------------------------------------------------------------------------------
        % CONFIGURATIONS
        %-------------------------------------------------------------------------------
        \documentclass[11pt, a4paper]{awesome-cv} 
        \geometry{left=1.4cm, top=.8cm, right=1.4cm, bottom=1.8cm, footskip=.5cm}
        \colorlet{awesome}{awesome-skyblue}
        \setbool{acvSectionColorHighlight}{true}
        \renewcommand{\acvHeaderSocialSep}{\quad \textbar\quad}
        %-------------------------------------------------------------------------------
        % PROFILE
        %-------------------------------------------------------------------------------
        """)
        # Assuming ProfileGenerator writes directly to the main_resume.tex
        profile_generator = ResumeGenerator(excel_file, main_resume_path)
        profile_generator.generate_tex()

        f.write(r"""
        \input{resume/summary.tex}
        \input{resume/skills.tex}
        \input{resume/competencies.tex}
        \input{resume/experience.tex}
        \input{resume/education.tex}

        \end{document}
        """)

    # Generate individual .tex files in the 'resume' folder
    SummaryGenerator(excel_file, os.path.join(resume_folder, 'summary.tex')).generate_tex()
    SkillsGenerator(excel_file, os.path.join(resume_folder, 'skills.tex')).generate_tex()
    CompetenciesGenerator(excel_file, os.path.join(resume_folder, 'competencies.tex')).generate_tex()
    ExperienceGenerator(excel_file, os.path.join(resume_folder, 'experience.tex')).generate_tex()
    EducationGenerator(excel_file, os.path.join(resume_folder, 'education.tex')).generate_tex()

if __name__ == "__main__":
    excel_file = "data.xlsx"  # Replace with the actual Excel file path
    main(excel_file)
