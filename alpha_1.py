
os.chdir(r"D:\un\diss\project")
def queue_text(seq):
    list_queue = list(range(1, x))
    for i in list_queue:
        if seq == 1:
            for root, dirs, files in os.walk('D:/un/diss/project/comp_science'):
                for file in files:
                    doc = docx.Document(r"D:/un/diss/project/comp_science/"+file)
                    y = 0
                    txt5 = ''
                    for paragraph in doc.paragraphs:
                        for run in paragraph.runs:
                            txt = run.text
                            txt2=txt.lower()
                            txt3=txt2.replace(",", "")
                            txt4 = txt3.replace(".", "")
                            txt5=txt4.strip()
                            if run.bold:
                                main_term = str(run.text)
                                main_term_name = "'"+main_term+"'"
                                query1 = "MERGE(t:Term{name:"+main_term_name+"})"
                                conn.query(query1, db='DB1')
                                receiver = main_term_name
                                instance_receiver = main_term_name
                            if run.italic and run.underline:
                                host_term = str(run.text)
                                host_term_name = "'"+host_term+"'"
                                query1 = "MATCH(t:Term{name:"+main_term_name+"}) MERGE(s:Term{name:"+host_term_name+"}) MERGE (t)-[:is_a]->(s)"
                                conn.query(query1, db='DB1')
                            if txt5 in ('англ'):
                                syn = 1
                                txt5 = ''
                            if run.italic and syn == 1:
                                sub_term = str(run.text)
                                sub_term_name = "'"+sub_term+"'"
                                query1 = "MATCH(t:Term{name:"+main_term_name+"}) MERGE(s:Term{name:"+sub_term_name+"}) MERGE (t)-[:synonym]->(s)"
                                conn.query(query1, db='DB1')
                                syn = 0
                            if run.italic and not run.underline:
                                sub_term = str(run.text)
                                sub_term_name = "'"+sub_term+"'"
                                query1 = "MATCH(t:Term{name:"+main_term_name+"}) MERGE(s:Term{name:"+sub_term_name+"}) MERGE (s)-[:is_part_of]->(t)"
                                conn.query(query1, db='DB1')
                                receiver = sub_term_name
                                instance_receiver = sub_term_name
                            if run.underline and not run.italic:
                                instance = str(run.text)
                                instance_name = "'"+instance+"'"
                                query1 = "MERGE(i:Instance{name:"+instance_name+"}) MATCH (t:Term{name:"+instance_receiver+"}) MERGE (i)-[:is_implementation]->(t)
                                conn.query(query1, db='DB1')

