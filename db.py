from sqlalchemy import create_engine, Column, Integer, String, LargeBinary, Boolean, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship

DATABASE_URI = 'sqlite:///macros.db'
Base = declarative_base()

class Document(Base):
    __tablename__ = 'document'
    id = Column(Integer, primary_key=True)
    name = Column(String(120), nullable=False)
    generated_pdf = Column(LargeBinary, nullable=True)
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

def save_document(name, pdf_data, macros):
    document = Document(name=name, generated_pdf=pdf_data)
    session.add(document)
    session.commit()

    for macro in macros:
        macro_record = Macro(name=macro['name'], document_id=document.id, efficient=macro.get('efficient', False), flowchart=macro.get('flowchart'))
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
