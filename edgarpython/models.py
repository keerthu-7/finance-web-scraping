from pydantic import BaseModel


class Submission(BaseModel):
    form: str
    accessionNumber: str
