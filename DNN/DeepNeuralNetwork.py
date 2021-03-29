import numpy as np
import pandas as pd
import tensorflow as tf
import xlwings
from sklearn.model_selection import train_test_split
from tensorflow import keras


def loading2Excel(path, shtname):
    wp = xlwings.App(visible=False, add_book=False)
    wb = wp.books.open(path)
    st = wb.sheets[shtname]
    merge = st.used_range.value
    wb.close()
    wp.quit()
    return merge


class dataLoader():
    def __init__(self):
        data = loading2Excel("./data/wholemat_n.xlsx", "wholemat_n")
        data = pd.DataFrame(data[1:], columns=data[0])
        data = data.iloc[:, 1:]
        mat = data.iloc[:, 1:]
        self.mat = mat.to_numpy()
        label = data.iloc[:, 0]
        self.label = label.to_numpy()
        self.train_data, self.test_data, self.train_y, self.test_y = train_test_split(mat, label, test_size=0.28)
        self.num_train_data = self.train_data.shape[0]

    def get_batch(self, batch_size):
        index = np.random.randint(0, self.num_train_data, batch_size)
        return self.train_data[index, :], self.train_y[index]

    def get_y(self):
        return pd.DataFrame(self.label)

    def get_X(self):
        return pd.DataFrame(self.mat)


def DNN_model():
    movieID = keras.layers.Input(shape=(1,), name="mID")
    movieHeat = keras.layers.Input(shape=(1,), name="mHeat")
    movieFeature = keras.layers.Input(shape=(22,), name="mFeature")
    criticID = keras.layers.Input(shape=(1,), name="cID")
    criticFeature = keras.layers.Input(shape=(22,), name="cFeature")
    movieVector = tf.keras.layers.concatenate([
        keras.layers.Embedding(1200, 128)(movieID),
        keras.layers.Embedding(574, 32)(movieHeat),
        keras.layers.Embedding(1200, 64)(movieFeature)
    ])
    movieVector = keras.layers.Dense(128, activation="relu")(movieVector)
    movieVector = keras.layers.Dense(64, activation="relu")(movieVector)
    movieVector = keras.layers.Dense(8, activation="relu", name="movieEmbedding", kernel_regularizer="12")(movieVector)
    criticVector = tf.keras.layers.concatenate([
        keras.layers.Embedding(4200, 256)(criticID),
        keras.layers.Embedding(4200, 256)(criticFeature)
    ])
    criticVector = keras.layers.Dense(128, activation="relu")(criticVector)
    criticVector = keras.layers.Dense(64, activation="relu")(criticVector)
    criticVector = keras.layers.Dense(8, activation="relu", name="criticEmbedding", kernel_regularizer="12")(
        criticVector)
    dotFeature = tf.reduce_sum(movieVector * criticVector, axis=1)
    dotFeature = tf.expand_dims(dotFeature, 1)
    result = keras.layers.Dense(1, activation="sigmoid")(dotFeature)
    return keras.models.Model(inputs=[movieID, movieHeat, movieFeature, criticID, criticFeature], outputs=[result])


