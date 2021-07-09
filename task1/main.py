from fastapi import FastAPI, HTTPException
from fastapi.requests import Request
import pandas as pd
from fastapi.templating import Jinja2Templates

app = FastAPI()
template = Jinja2Templates(directory="templates")

class Recommender:
    def __init__(self, path=""):
        self.df = pd.read_csv(path+"output.csv")
        print("found output.csv")

    def recommend(self, isbn):
        row = self.df.loc[self.df["isbn"]==isbn, ["class_labels", "title"]]
        if row["class_labels"].size==1:
            title = row["title"].values[0]
            class_label = row["class_labels"].values[0]
            cluster_movies = self.df.loc[(self.df["class_labels"]==class_label)&(self.df["isbn"]!=isbn), "title"].values[:10]
            return (title, list(cluster_movies))
        else:
            return None

recommender  = Recommender()

@app.get("/")
def index(request: Request):
    return template.TemplateResponse("index.html", {"request": request})

@app.get("/books")
def movies_list(isbn: str):
    response = recommender.recommend(isbn)
    if response is not None:
        title = response[0]
        movies_list = response[1]
        return {"title": title, "movies_list": movies_list}
    else:
        raise HTTPException(404, detail="isbn not found")