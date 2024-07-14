from sqlalchemy import create_engine, Column, Integer, String, LargeBinary, Boolean, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship

DATABASE_URI = 'sqlite:///macros.db'
Base = declarative_base()

class Document(Base):
    __tablename__ = 'document'
    id = Column(Integer, primary_key=True)
    name = Column(String(120), nullable=False)
    functional_pdf = Column(LargeBinary, nullable=True)
    analysis_pdf = Column(LargeBinary, nullable=True)
    macros = relationship('Macro', backref='document', lazy=True)

class Macro(Base):
    __tablename__ = 'macro'
    id = Column(Integer, primary_key=True)
    name = Column(String(120), nullable=False)
    document_id = Column(Integer, ForeignKey('document.id'), nullable=False)
    flowchart = Column(LargeBinary, nullable=True)
    efficient = Column(Boolean, default=False)

engine = create_engine(DATABASE_URI)
Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
session = Session()

def save_document(name, functional_pdf_data, analysis_pdf_data, macros, logic_explanations):
    document = Document(name=name, functional_pdf=functional_pdf_data, analysis_pdf=analysis_pdf_data)
    session.add(document)
    session.commit()

    for idx, macro in enumerate(macros):
        # Read flowchart file as bytes
        flowchart_path = logic_explanations[idx]['process_flowchart'] if idx < len(logic_explanations) else None
        flowchart_bytes = None
        if flowchart_path:
            with open(flowchart_path, 'rb') as f:
                flowchart_bytes = f.read()

        # Insert macro record into database
        macro_record = Macro(
            name=macro['name'],
            document_id=document.id,
            efficient=macro.get('efficient', False),
            flowchart=flowchart_bytes
        )
        session.add(macro_record)

    session.commit()
    return document.id

def get_all_documents():
    return session.query(Document).all()

def get_document_by_id(document_id):
    return session.query(Document).filter(Document.id == document_id).first()

def get_all_macros():
    return session.query(Macro).all()

def get_macros_by_document_id(document_id):
    return session.query(Macro).filter(Macro.document_id == document_id).all()

def get_macro_by_id(macro_id):
    return session.query(Macro).filter(Macro.id == macro_id).first()
