import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.neural_network import MLPClassifier
from sklearn.model_selection import train_test_split
from sklearn import metrics
import matplotlib as plt
scaler = StandardScaler()
df = pd.read_csv('data2/testing.csv')

label = ['N']
features = ['C','D','E','F','H','I','J','K']
df[features] = scaler.fit_transform(df[features])

X = df[features].values
y = df[label].values
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.30, random_state=40)

clf = MLPClassifier(hidden_layer_sizes=(50,20,10), max_iter=500, alpha=0.0001,activation = 'relu',solver='adam',random_state=1)
clf.fit(X_train,y_train.ravel())
y_pred = clf.predict(X_test)
print("Accuracy:",metrics.accuracy_score(y_test, y_pred))
cm = metrics.confusion_matrix(y_test,y_pred)
print("Confusion matrix:\n",cm)
p = cm[1][1]/(cm[1][1]+cm[0][1])
r = cm[1][1]/(cm[1][1]+cm[1][0])
print("precision: ",p)
print("recall: ",r)
print("F Score",(2*p*r)/(p+r))



