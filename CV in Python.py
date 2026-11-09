import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import urllib.request
import io
from PIL import Image

def create_element(name):
    return OxmlElement(name)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets the internal indents in table cells for precise alignment."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = create_element('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = create_element(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_p_border_bottom(paragraph, color_hex="1B365D", size="6"):
    """Adds a solid horizontal line below the section heading."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = create_element('w:pBdr')
    bottom = create_element('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)

def build_resume():
    doc = Document()
    
    # 1. Setting the margins (ATS Standard—0.75 inches to maximize usable space)
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        
    # 2. Basic Font Styles (Classic, Clean Calibri - surely not Arial:)
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    
    NAVY = RGBColor(0x1B, 0x36, 0x5D)
    
        
     # --- HEADER: TABLE WITH PHOTOS AND CONTACT INFORMATION ---
    # I need to create an invisible table with 1 row and 2 columns to neatly align the photo on the right
    header_table = doc.add_table(rows=1, cols=2)
    header_table.autofit = False
    header_table.columns[0].width = Inches(5.4) # Текст
    header_table.columns[1].width = Inches(1.6) # Фото
    
    hc0, hc1 = header_table.cell(0, 0), header_table.cell(0, 1)
    set_cell_margins(hc0, top=0, bottom=0, left=0, right=0)
    set_cell_margins(hc1, top=0, bottom=0, left=0, right=0)
    
    # Left column: Name and contact information
    p_name = hc0.paragraphs[0]
    p_name.paragraph_format.space_after = Pt(2)
    run_name = p_name.add_run("ALEXANDER SHEMETEV")
    run_name.font.size = Pt(22)
    run_name.font.bold = True
    run_name.font.color.rgb = NAVY
    
    p_meta = hc0.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(0)
    meta_text = "Yerevan, Armenia | +374-55-77-5786 | alexandershemetev@gmail.com\nLinkedIn: https://www.linkedin.com/in/alexshemetev/"
    run_meta = p_meta.add_run(meta_text)
    run_meta.font.size = Pt(10)
    
    # Right cell: Photo 
    p_photo = hc1.paragraphs[0]
    p_photo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    img_stream = None
    
    # Attempt 1: Downloading via URL
    try:
        url = "https://samlib.ru/s/shemetew_a_a/.photo1.jpg"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            img_stream = io.BytesIO(response.read())
    except Exception as e:
        print(f"URL photo download failed: {e}")
        
    # Attempt 2: If the network goes down, use the local file
    if img_stream is None:
        photo_path = r"E:\Pictures\Alex Shemetev.jpg"
        if os.path.exists(photo_path):
            with open(photo_path, 'rb') as f:
                img_stream = io.BytesIO(f.read())
                
    # Stream Processing and Handling ZeroDivisionErrors potential problems (avoiding a famous bug)
    if img_stream is not None:
        try:
            # Open the image using PIL to check the metadata
            with Image.open(img_stream) as img:
                # If the DPI is missing or corrupted, force the standard 96 DPI
                if not img.info.get('jpeg_dpi') or img.info.get('jpeg_dpi')[0] == 0:
                    output_stream = io.BytesIO()
                    img.save(output_stream, format='JPEG', dpi=(96, 96))
                    output_stream.seek(0)
                    img_stream = output_stream
                else:
                    img_stream.seek(0)
            
            # Insert a pre-validated image without the risk of division by zero
            p_photo.add_run().add_picture(img_stream, height=Inches(1.8))
        except Exception as img_err:
            print(f"Image processing failed: {img_err}")
            p_photo.add_run("[ Photo ]")
    else:
        p_photo.add_run("[ Photo ]")
    
    
    
    def add_section_heading(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text.upper())
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = NAVY
        add_p_border_bottom(p)
        return p

    # --- CHAPTER 1: PROFESSIONAL SUMMARY ---
    add_section_heading("Professional Summary")
    #p_sum = doc.add_paragraph()
    #p_sum.paragraph_format.space_after = Pt(6)
    #p_sum.paragraph_format.line_spacing = 1.15
    p_sum = doc.add_paragraph()
    p_sum.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_sum.paragraph_format.space_after = Pt(2)
    p_sum.paragraph_format.line_spacing = 1.05
    
    p_sum.add_run(
        "Analytically rigorous and results-driven Strategy & Finance professional with a Ph.D. equivalent "
        "in Economics and extensive global experience in financial modeling, corporate risk assessment, "
        "and quantitative analysis. Expert in translating complex data arrays into clear, actionable insights "
        "for C-suite executives and board-level decision-making. Proven track record in orchestrating cross-functional "
        "digital transformation initiatives, evaluating socio-economic resilience, and establishing scalable baseline "
        "metrics for performance tracking. Fluent in English and Russian, combining strong strategic insight with an "
        "advanced technological background to optimize resource allocation, evaluate new business ventures, and drive "
        "organizational growth."
    )
    
    # --- CHAPTER 2: CORE COMPETENCIES ---
    add_section_heading("Core Competencies")
    comp_data = [
        ("Corporate & Growth Strategy", "Strategic Planning, Market Trends Analysis, Competitor Assessment, Growth Mapping, Business Planning"),
        ("Quantitative Finance & Analysis", "Advanced Financial Modeling, Risk Forecasting (Proprietary A-Matrix Framework), Big Data Analytics (R, SQL, STATA)"),
        ("Performance Management", "Cascading Strategic Goals, OKR/KPI Implementation, Process Automation, Cross-functional Alignment"),
        ("Executive Communication", "C-Suite Presentations, Data-driven Investment Briefs, Strategic Reports, Board Materials Preparation")
    ]
    for category, skills in comp_data:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(2)
        r_cat = p.add_run(f"{category}: ")
        r_cat.bold = True
        p.add_run(skills)

    # --- CHAPTER 3: PROFESSIONAL EXPERIENCE ---
    add_section_heading("Professional Experience")
    
    jobs = [
        {
            "title": "Strategic Advisor & Software Engineer",
            "company": "Private Practice / FinSoft.Systems",
            "date": "03/2012 – Present",
            "bullets": [
                "Engineered and commercialized advanced quantitative financial modeling software used to assess market opportunities and evaluate high-level corporate risks.",
                "Developed a unique proprietary predictive algorithm (A-Matrix framework) capable of accurately estimating company health and missing line items under conditions of data scarcity.",
                "Authored 12 strategic and analytical textbooks on corporate finance, economic metrics, and risk management, achieving a readership of over 2 million globally via distribution pipelines.",
                "Automated end-to-end data analytics and trend visualization frameworks to accelerate decision-making timelines for external corporate clients."
            ]
        },
        {
            "title": "Project Director & Lead Advanced Trainer",
            "company": "North-West Institute of Management",
            "date": "09/2014 – 07/2022",
            "bullets": [
                "Directed the conceptualization and software rollout of a key automated regional development platform ('Digital Transformation for Heads of Departments') to evaluate systemic risk maps and resource alignment.",
                "Led strategic planning and program design for comprehensive data skills optimization under national development initiatives ('Demography' program).",
                "Managed large cross-functional teams and operational workstreams; twice honored with top university awards for structural innovation and program excellence."
            ]
        },
        {
            "title": "Strategic Project Leader & Ph.D. Researcher",
            "company": "Prague University of Economics and Business (VŠE)",
            "date": "02/2020 – Present",
            "bullets": [
                "Spearheaded a dedicated cross-functional analytics project evaluating high-frequency pandemic market impacts and consumer resilience, managing a direct institutional budget of 150,000 CZK.",
                "Co-led an international comparative expansion assessment across Spain, Finland, and Estonia to design data-driven competitive strategies and policy recommendations for the Czech Republic (Budget: 670,000 CZK).",
                "Established and optimized high-stakes communication protocols with foreign entities and partner research teams (e.g., University of Turku) for complex international grants."
            ]
        },
        {
            "title": "Startup Founder & Entrepreneur",
            "company": "Venture Portfolio (Real Estate, Hospitality, Medicine)",
            "date": "09/2009 – Present",
            "bullets": [
                "Launched and scaled multiple niche ventures; directly oversaw capital allocation, high-stakes negotiations, cash-flow modeling, and end-to-end long-term growth strategies."
            ]
        }
    ]
    for job in jobs:
        # 1. Create a table with a rigid, fixed layout (Layout: Fixed)
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        
        # Set a fixed column width distribution at the table level
        tblPr = table._tbl.tblPr
        tblLayout = create_element('w:tblLayout')
        tblLayout.set(qn('w:type'), 'fixed')
        tblPr.append(tblLayout)
        
        c0, c1 = table.cell(0, 0), table.cell(0, 1)
        
        # 2. I explicitly set the cell widths in the XML (14.0 cm for text and 3.8 cm for dates)        tcPr0 = SETTING IT!

        tcPr0 = c0._tc.get_or_add_tcPr()
        tcW0 = create_element('w:tcW')
        tcW0.set(qn('w:w'), str(int(14.0 * 567))) # 14 cm
        tcW0.set(qn('w:type'), 'dxa')
        tcPr0.append(tcW0)
        
        tcPr1 = c1._tc.get_or_add_tcPr()
        tcW1 = create_element('w:tcW')
        tcW1.set(qn('w:w'), str(int(3.8 * 567)))  # 3.8 cm
        tcW1.set(qn('w:type'), 'dxa')
        tcPr1.append(tcW1)
        
        # 3. Use `noWrap` and set indents to zero for maximum capacity
        tcPr0.append(create_element('w:noWrap'))
        tcPr1.append(create_element('w:noWrap'))
        set_cell_margins(c0, top=40, bottom=0, left=0, right=0)
        set_cell_margins(c1, top=40, bottom=0, left=0, right=0)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        p0.paragraph_format.line_spacing = 1.0
        
        r_title = p0.add_run(job["title"])
        r_title.bold = True
        p0.add_run(f" | {job['company']}")
        
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = 1.0
        
        r_date = p1.add_run(job["date"])
        r_date.bold = True
        
        for b in job["bullets"]:
            pb = doc.add_paragraph(style='List Bullet')
            pb.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pb.paragraph_format.space_after = Pt(2)
            pb.paragraph_format.line_spacing = 1.05
            pb.add_run(b)
 
  
  # --- INSERT KEY STRATEGIC PROJECTS SECTION ---
    add_section_heading("Key Strategic Projects")
    projects = [
        {
            "name": "Czech Pharmaceutical Market Decadal Big Data Optimization",
            "desc": "Designed an integrated R and SQL analytics framework processing a decade of clinical and commercial data (2010-2020). Modeled the systemic impact of VZP and SZP health insurance frameworks on corporate revenue structures to support capital allocation strategies."
        },
        {
            "name": "Smartphone Geopositioning Consumer Behavior Analytics",
            "desc": "Built machine learning workflows in R to classify minute behavior shifts using multidimensional physical smartphone location streams. Developed high-precision commercialization algorithms used to target hyper-personalized marketing ads dynamically."
        },
        {
            "name": "Hurricane Lifecycles High-Frequency Vector Visualization",
            "desc": "Engineered customized dynamic coordinate mapping code to map real-time vector paths, atmospheric pressures, and wind velocities across multi-stage trajectories."
        }
    ]
    for proj in projects:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.keep_with_next = True
        r_pname = p.add_run(proj["name"])
        r_pname.bold = True
        
        pd = doc.add_paragraph()
        pd.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pd.paragraph_format.space_after = Pt(2)
        pd.paragraph_format.line_spacing = 1.05
        pd.add_run(proj["desc"])


  # --- CHAPTER 4: EDUCATION ---
    add_section_heading("Education")
    edu_items = [
        ("Ph.D. equivalent (Candidate of Sciences in Economics & Finance)", "Saint-Petersburg State University of Economics", "Grade: 100% | Focus: Quantitative Risk Estimation"),
        ("Master’s Degree in Applied Economics", "CERGE-EI (Joint workplace of Charles University & Czech Academy of Sciences)", "Grade: 90% | US-Accredited Curriculum, Advanced Econometrics"),
        ("Master of Business Administration (MBA) in Crisis Management", "Ural State University", "Grade: 100%"),
        ("Master’s Degree in Corporate, Banking, and Securities Law", "Ural State Agricultural Academy", "Grade: 100%"),
        ("Master’s Degree in English Language and Literature", "Ural State University", "Grade: 100%")
    ]
    
    for degree, school, detail in edu_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        r_deg = p.add_run(degree)
        r_deg.bold = True
        p.add_run(f" – {school} ({detail})")
        
    # --- CHAPTER 5: TECHNICAL SKILLS & CERTIFICATIONS ---
    add_section_heading("Technical Skills & Certifications")
    
    # Paragraph 1: Development Tools and Languages
    p_tech = doc.add_paragraph()
    p_tech.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_tech.paragraph_format.space_after = Pt(2)
    p_tech.paragraph_format.line_spacing = 1.05
    
    r_tools_label = p_tech.add_run("Quantitative Tools & Languages: ")
    r_tools_label.bold = True
    p_tech.add_run("R (Advanced Frameworks), SQL, STATA, EViews, Gretl, HTML, Shiny + Plotly web automation, C# (Microsoft Certified).")
    
    # Paragraph 2: Key certificates (created using the full Enter/add_paragraph command)
    p_cert = doc.add_paragraph()
    p_cert.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cert.paragraph_format.space_after = Pt(2)
    p_cert.paragraph_format.line_spacing = 1.05
    
    r_cert_label = p_cert.add_run("Key Certifications: ")
    r_cert_label.bold = True
    p_cert.add_run(
        "Certified Quantitative Research Methods Specialist (VŠE, 2022) | "
        "Data Analytics for Business (University of Colorado Boulder) | "
        "Visualization for Data Journalism (University of Illinois) | "
        "CISCO Certified System Administrator (CCNA)."
    )
    
    # --- CHAPTER 6: LANGUAGES ---
    add_section_heading("Languages")
    p_lang = doc.add_paragraph()
    p_lang.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_lang.paragraph_format.space_after = Pt(2)
    p_lang.paragraph_format.line_spacing = 1.05
    p_lang.add_run("Russian (Native) | English (Fluent / Professional Mastery) | Czech (Fluent) | German (Intermediate) | French (Intermediate)")
    
     # --- INSERTING THE HOBBIES SECTION ---
    add_section_heading("Hobbies")
    p_hob = doc.add_paragraph()
    p_hob.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_hob.paragraph_format.space_after = Pt(2)
    p_hob.paragraph_format.line_spacing = 1.05
    p_hob.add_run("Classical Music, Jazz, Theatres, Reading, Walking")
    
    
    
    
       # --- SAVE FILE VIA RELATIVE PATHWAYS FOR INDEPENDENT RUNS ---
    target_directory = "output"
    os.makedirs(target_directory, exist_ok=True)
    
    output_filename = os.path.join(target_directory, "Alexander_Shemetev_Resume_Strategy_Associate.docx")
    doc.save(output_filename)
    print(f"Success! ATS Document successfully rendered at: {output_filename}")

if __name__ == '__main__':
    build_resume()
    