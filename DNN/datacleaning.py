import xlwings
import pandas as pd
import tensorflow as tf
import numpy as np
from sklearn.model_selection import train_test_split

CriticPath="./data/critic.reviews.normalised.csv"
MoviePath="./data/moviedata1021.xlsx"

def loadingCritics(Path):
    critics = pd.read_csv(Path)
    critics.head(10)
    critics.columns
    outdf=pd.DataFrame()
    outdf["mLink"]=critics["Title"]
    outdf["Reviewer"]=critics["Reviewer"]
    outdf["mScore"]=critics["norm.review.score"]
    outdf.head(10)
    print(outdf.isnull().any())
    print(outdf.isnull().sum())
    outdf["Reviewer"].fillna("Carl Johnson",inplace=True)
    outdf["mScore"].fillna(0,inplace=True)
    return outdf
    print("loading Critics complete")

def loadingMovies(path,sheet):
    wp = xlwings.App(visible=False, add_book=False)
    wb=wp.books.open(path)
    st=wb.sheets[sheet]
    content=st.used_range.value
    wb.close()
    wp.quit()
    content=pd.DataFrame(content)
    outdf=pd.DataFrame()
    outdf["mName"]=content[0]
    outdf["uHeat"]=content[8]
    outdf["Genres"]=content[11]
    outdf["mLink"]=content[23]
    outdf.isnull().sum()
    outdf = outdf.dropna()
    return outdf

def mergingDatasets(Cdf,Mdf):
    mergedf=pd.merge(Cdf,Mdf,how="inner",on="mLink")
    print("Finish Merging")
    # LOADING
    mergedf=mergedf[~(mergedf["mName"]=="NA")]
    mergedf=mergedf[~(mergedf["Genres"]=="NA")]
    mergedf["uHeat"].replace("NA","0",inplace=True)
    mergedf["uHeat"]=mergedf["uHeat"].astype("float64")
    mergedf["uHeat"][mergedf["uHeat"] == 0] = mergedf["uHeat"].mean()
    mergedf["Genres"]=mergedf["Genres"].apply(replaceStr)
    namedf=pd.DataFrame()
    namedf["mName"]=mergedf["mName"].unique()
    namedf["mID"]=namedf.index
    mergedf=pd.merge(mergedf,namedf,how="left",on="mName")
    criticdf=pd.DataFrame()
    criticdf["Reviewer"]=mergedf["Reviewer"].unique()
    criticdf["cID"] = criticdf.index
    mergedf = pd.merge(mergedf, criticdf, how="left", on="Reviewer")
    mergedf = mergedf.drop_duplicates()
    # SAVE
    # save2Excel(mergedf, "merged", "./data/merged.xlsx")
    return mergedf

def replaceStr(text:str):
    text=text.replace("Anime","Animation")
    return text

def save2Excel(df,shtname,path):
    wb=xlwings.Book()
    sht=wb.sheets[0]
    sht.name=shtname
    sht.range("A1").value=df
    wb.save(path)
    xlwings.App().quit()

def loading2Excel(path,shtname):
    wp = xlwings.App(visible=False, add_book=False)
    wb = wp.books.open(path)
    st = wb.sheets[shtname]
    merge = st.used_range.value
    wb.close()
    wp.quit()
    return merge

def createGenrelist(l:list):
    genlist=[]
    for r in l:
        genre=r.replace(" And ",", ")
        genres=genre.split(", ")
        for g in genres:
            if g in genlist:
                continue
            else:
                genlist.append(g)
    return genlist

def rowlist(getext):
    genlist=[]
    genre=getext.replace(" And ",", ")
    genres=genre.split(", ")
    for g in genres:
        if g in genlist:
            continue
        else:
            genlist.append(g)
    return genlist

def rowlistRepeat(getext):
    genlist=[]
    genre=getext.replace(" And ",", ")
    genres=genre.split(", ")
    for g in genres:
        genlist.append(g)
    return genlist

def movieEncoding(df,gelist):
    movie=pd.DataFrame()
    movie["mName"]=df["mName"].unique()
    movie=pd.merge(movie,df,how="left",on="mName")
    movie["Genres"]=movie["Genres"].apply(rowlist)
    for g in gelist:
        movie[g]=0
    for index, row in movie.iterrows():
        for g in row["Genres"]:
            if g in gelist:
                movie.loc[index,g] = 1
    # save2Excel(movie,"moviemat","./data/moviemat.xlsx")
    return movie

def criticEncoding(df,gelist):
    criticdf=pd.DataFrame()
    criticdf["Reviewer"]=df["Reviewer"].unique()
    for g in gelist:
        criticdf[g]=0
    for index,c in criticdf.iterrows():
        critic_movie=df[df["Reviewer"]==c["Reviewer"]]
        genre=""
        for ind1,r in critic_movie.iterrows():
            genre += r["Genres"] + ", "
        genre=genre[:-2]
        genre=genre.replace(" And ", ", ")
        genre=genre.split(", ")
        for g in genre:
            if g in gelist:
                criticdf.loc[index,g] += 1
    save2Excel(criticdf, "criticmat", "./data/criticmat.xlsx")
    return criticdf

def listRename(l:list,suffix:str):
    out={}
    for i in l:
        t=i+suffix
        out[i]=t
    return out

def makingWhole():
    movie=loading2Excel("./data/merged.xlsx","merged")
    movie = pd.DataFrame(movie[1:], columns=movie[0])
    mFeature=loading2Excel("./data/moviemat.xlsx","moviemat")
    mFeature = pd.DataFrame(mFeature[1:], columns=mFeature[0])
    mFeature.drop(columns=[None, "mName", "mLink", "Reviewer", "Genres", "cID"], inplace=True)
    names = listRename(mFeature.columns, "_m")
    mFeature.rename(columns=names,inplace=True)
    mFeature.drop(columns=["mID_m", 'mScore_m', 'uHeat_m'], inplace=True)
    cFeature=loading2Excel("./data/criticmat.xlsx","criticmat")
    # cFeature = pd.DataFrame(cFeature[1:], columns=cFeature[0])
    # cFeature=cFeature.iloc[:,1:]
    # names = listRename(cFeature.columns, "_c")
    # cFeature.rename(columns=names,inplace=True)
    # cFeature["Reviewer"]=cFeature["Reviewer_c"]
    # cFeature.drop(columns=["Reviewer_c"], inplace=True)
    outdf = movie[['mScore', 'uHeat', 'mID', 'cID', "Reviewer"]]
    criticlist=outdf["Reviewer"].to_list()
    clist=[]
    for c in criticlist:
        l=[]
        l.append(c)
        for f in cFeature:
            if c==f[1]:
                l.append(f[2:])
        clist.append(l)
    clist2=[]
    for c in clist:
        c=str(c)
        c=c.replace("[","")
        c = c.replace("]", "")
        c=list(eval(c))
        clist2.append(c)
    cFeature=pd.DataFrame(clist2,columns=cFeature[0][1:])
    names = listRename(cFeature.columns, "_c")
    cFeature.rename(columns=names, inplace=True)
    cFeature.drop(columns=["Reviewer_c"], inplace=True)
    outdf = pd.concat([outdf, mFeature, cFeature], axis=1)
    outdf.drop(columns=["Reviewer"], inplace=True)
    # save2Excel(outdf, "wholemat", "./data/wholemat.xlsx")
    outdf_m = (outdf - outdf.min()) / (outdf.max() - outdf.min())
    save2Excel(outdf_m,"wholemat_n", "./data/wholemat_n.xlsx")
    return outdf

    # cFeature
    # [[None, 'Reviewer', 'Action', 'Adventure', 'Comedy', 'Fantasy', 'Mystery', 'Thriller', 'Drama', 'Sci Fi', 'Kids',
    #   'Family', 'Animation', 'Horror', 'Musical', 'Western', 'Romance', 'Crime', 'War', 'History', 'Documentary',
    #   'Sports', 'Fitness', 'Other'],
    #  [0.0, 'Kelechi Ehenulo', 30.0, 29.0, 7.0, 21.0, 8.0, 8.0, 5.0, 32.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 1.0, 0.0, 0.0,
    #   0.0, 0.0, 0.0, 0.0, 0.0],

def creatingTables():
    criticdf=loadingCritics(CriticPath)
    moviedf=loadingMovies(MoviePath,"Movie")
    merged=mergingDatasets(criticdf,moviedf)
    gelist=createGenrelist(merged["Genres"].tolist())
    moviedf=movieEncoding(merged,gelist)
    criticdf=criticEncoding(merged,gelist)
    print("Movie, mFeature and cFeature are created!")

class dataLoader():
    def __init__(self):
        data=loading2Excel("./data/wholemat_n.xlsx","wholemat_n")
        data = pd.DataFrame(data[1:], columns=data[0])
        data = data.iloc[:, 1:]
        mat=data.iloc[:,1:]
        mat=mat.to_numpy()
        label=data.iloc[:,0]
        label=label.to_numpy()
        self.train_data,self.test_data,self.train_y,self.test_y=train_test_split(mat,label,test_size=0.28)
        self.num_train_data=self.train_data.shape[0]
    def get_batch(self,batch_size):
        index=np.random.randint(0,self.num_train_data,batch_size)
        return self.train_data[index,:],self.train_y[index]

class DNN(tf.keras.Model):
    def __init__(self):
        super().__init__()
        self.dense1=tf.keras.layers.Dense(units=100,activation=tf.nn.relu)
        self.dense2 = tf.keras.layers.Dense(units=80, activation=tf.nn.relu)
        self.dense3 = tf.keras.layers.Dense(units=60, activation=tf.nn.relu)
        self.dense4=tf.keras.layers.Dense(units=1,activation=tf.nn.sigmoid)

    def call(self,inputs):
        x=self.dense1(inputs)
        x=self.dense2(x)
        x=self.dense3(x)
        output=self.dense4(x)
        return output

def runDNN():
    num_epoch = 5
    batch_size = 136
    learning_rate = 0.0001
    model = DNN()
    data_loader = dataLoader()
    optimizer = tf.keras.optimizers.Adam(learning_rate=learning_rate)
    #optimizer = tf.keras.optimizers.SGD(learning_rate=learning_rate)
    num_batches = int(data_loader.num_train_data // batch_size * num_epoch)
    for batch_index in range(num_batches):  # num_batches
        X, y = data_loader.get_batch(batch_size)
        with tf.GradientTape() as tape:
            y_pred = model(X)
            loss = tf.reduce_sum(tf.square(y_pred-y))
            print("batch %d: loss: %f" % (batch_index, loss.numpy()))
        grads = tape.gradient(loss, model.variables)
        optimizer.apply_gradients(grads_and_vars=zip(grads, model.variables))
    train_loss=tf.keras.metrics.Mean("train_loss",dtype=tf.float32)
    train_accuracy=tf.keras.metrics.MSE("train_accuracy")
    test_loss=tf.keras.metrics.Mean("test_loss",dtype=tf.float32)
    test_accuracy=tf.keras.metrics.MSE("test_accuracy")


if __name__=="__main__":
    # creatingTables()
    # makingWhole()
    print("Hello world")
    runDNN()


