# -*- coding: utf-8 -*-
from sqlalchemy import Column, String, create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base


class ALMSSQL():
    def __init__(self):
        self.engine = create_engine("mssql+pymssql://sa:~ljm56522319@192.165.4.58:1433/YGDB_C")
        self.DBSession = sessionmaker(bind=self.engine)
        self.session = self.DBSession()

    def ExecQuery(self, sql):
        result = self.session.execute(sql).fetchall()
        self.session.close()
        return result

    def ExecNonQuery(self, sql):
        self.session.execute(sql)
        self.session.flush()
        self.session.commit()
        self.session.close()

