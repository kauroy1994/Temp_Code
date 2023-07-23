import pandas as pd
from rdflib import Graph, Literal, URIRef
from rdflib.namespace import OWL, RDF, RDFS, XSD
from os import system

class FileOps(object):
    
    @staticmethod
    def read_xlsx(file_name):

        return pd.read_excel(file_name)
    
    @staticmethod
    def df_column_names_as_list(df):

        return list(df.columns)
    
class StringOps(object):

    @staticmethod
    def has_common_sub_string(str1,str2):

        if str1 in str2:
            return True
        
        return False
    
class CaseKnowledgeGraph(Graph):

    def __init__(self,name_space = None):

        super().__init__()
        self.name_space = name_space
        self.edges = []

    def add_case_resource(self,df):
        self.case_df = df

    def add_edge(self, edge):

        if edge not in self.edges:
            self.add(edge)
            self.edges.append(edge)

    def remove_edge(self, edge):

        if edge in self.edges:
            self.remove(edge)
            self.edges.remove(edge)

    def create_class(self,class_name=None):

        if class_name is None:
            print ('no class name found')
            exit()

        class_URI = URIRef(self.name_space+'#'+class_name)

        #add as OWL class to maintain standard
        edge = (class_URI,RDFS.subClassOf,OWL.Thing)
        self.add_edge(edge)

    def create_instance(self,instance_id,
                        class_name=None):

        if class_name is None:
            print ('no class name found')
            exit()

        class_URI = URIRef(self.name_space+'#'+class_name)

        #0th column assumed id
        i = instance_id
        instance_URI = URIRef(self.name_space+'#'
                              +class_name+str(self.case_df.loc[i,FileOps.df_column_names_as_list(self.case_df)[0]]))
        
        edge = (instance_URI,RDF.type,class_URI)
        self.add_edge(edge)
        return instance_URI
    
    def create_reln(self,reln_name = None):

        if reln_name is None:
            print ('relationship name not found')
            exit()

        reln_URI = URIRef(self.name_space+'#'+reln_name)
        edge = (reln_URI,RDFS.subPropertyOf,OWL.AnnotationProperty)
        self.add_edge(edge)
        return reln_URI
    
    def write_kg(self,
                 name):
        
        system('rm '+name+'.ttl')
        self.serialize(destination=name+'.ttl')

    def populate_kg_from_cases_df(self,df):

        #add resource
        self.add_case_resource(df)

        #get df column names
        df_column_names = FileOps.df_column_names_as_list(self.case_df)

        #create case class
        self.create_class('Case')

        #iterate through case files
        df_range = len(self.case_df) #precompute for efficiency

        for i in range(df_range):
            instance_URI = self.create_instance(i,'Case')
            for col in df_column_names:

                #ignore columns
                if True in [StringOps.has_common_sub_string(x,col) for x in ['Unnamed','Content','PDF']]:
                    continue

                #handle case Title
                if 'Title' in col:
                    if not str(self.case_df.Title[i]) == 'nan':
                        reln_URI = self.create_reln('Title')
                        edge = (instance_URI,
                                reln_URI,
                                Literal(str(self.case_df.Title[i])))
                        self.add_edge(edge); continue
    

                #handle headings
                if StringOps.has_common_sub_string('Heading',col):
                    heading_no = col.split(' ')[-1]
                    heading = str(self.case_df.loc[i,col])
                    content = self.case_df.loc[i,"Content "+heading_no]

                    #skip nan values
                    if str(heading) == 'nan' or str(content) == 'nan':
                        continue

                    heading_reln_URI = self.create_reln(heading.replace(' ','_'))
                    edge = (instance_URI,
                            heading_reln_URI,
                            Literal(content))
                    self.add_edge(edge)
                    continue
                
                
                #handle other non-ignored columns
                try:
                    col_value = self.case_df[i,col]
                except:
                    if 'Type' in col: #special case
                        col_value = self.case_df.Type[i] 

                #skip NaN values
                if str(col_value) == 'nan':
                    continue
                col_reln_URI = self.create_reln(col.replace(' ','_'))
                edge = (instance_URI,
                        col_reln_URI,
                        Literal(col_value))
                self.add_edge(edge)
                continue
    
if __name__ == '__main__':

    xlsx_file = "CaseStudy_Excel.xlsx"
    df = FileOps.read_xlsx(xlsx_file)
    name_space = "https://myCompany.com/"
    kg = CaseKnowledgeGraph(name_space = name_space)
    kg.populate_kg_from_cases_df(df)
    kg.write_kg(name="Test_Build_Ontology")
