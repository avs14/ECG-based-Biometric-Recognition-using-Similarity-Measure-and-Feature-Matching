import tensorflow
import pandas as pd
import numpy as np
import tensorflow as tf
from sklearn.preprocessing import StandardScaler
from tensorflow import keras as k
from sklearn.model_selection import train_test_split
from matplotlib import pyplot as plt
import time

start = time.time()
scaler = StandardScaler()
df = pd.read_csv('data2/testing.csv')
label = ['N']
features = ['C','D','E','F','H','I','J','K','M']
df[features] = scaler.fit_transform(df[features])
X = df[features].values
y = df[label].values
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.30, random_state=40)
n = X_train.shape[1]
model = k.Sequential()
model.add(k.layers.Dense(50, activation='relu',kernel_initializer ='he_normal', input_shape= (n,)) )
model.add(k.layers.Dense(20, activation='relu', kernel_initializer ='he_normal'))
model.add(k.layers.Dense(10, activation='relu', kernel_initializer ='he_normal'))
model.add(k.layers.Dense(1, activation='sigmoid'))

model.compile(optimizer = 'adam', loss = 'binary_crossentropy', metrics = ['accuracy'])
#tf_callbacks = k.callbacks.TensorBoard(log_dir = "logs/fit" , histogram_freq = 1)
history = model.fit(X_train, y_train, validation_data =(X_test,y_test), epochs=25, batch_size =32, verbose=1 )
loss, acc = model.evaluate(X_test,y_test, verbose =1)
delta = time.time()-start
print("time: ",delta)
pd.DataFrame(history.history).plot(figsize=(10,7))

print("accuracy and loss:\n",acc,"\n",loss)


plt.xlabel('Epochs')
plt.ylabel('Accuracy/Loss')
plt.title('Training Accuracy and Loss on ECG dataset')
plt.show()
plt.legend(history.history)





