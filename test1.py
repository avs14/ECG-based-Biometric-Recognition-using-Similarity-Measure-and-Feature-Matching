import pandas as pd
from sklearn.tree import DecisionTreeClassifier # Import Decision Tree Classifier
from sklearn.model_selection import train_test_split # Import train_test_split function
from sklearn import metrics
from sklearn.preprocessing import StandardScaler
import numpy as np

col = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N']

pima = pd.read_csv(r'data2/testing2.csv',header=None, names=col)
pima2 = pd.read_csv(r'data2/testing1.csv',header=None, names=col)
pima.head()
pima2.head()
scaler = StandardScaler()
features = ['C','D','E','F','H','I','J','K','M']
pima[features] = scaler.fit_transform(pima[features])
X = pima[features]
y = pima.N

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=40)

clf = DecisionTreeClassifier(criterion="entropy", max_depth=3)
clf = clf.fit(X.values,y.values)
y_pred = clf.predict(X_test.values)




print("Accuracy:",metrics.accuracy_score(y_test, y_pred))
cm = metrics.confusion_matrix(y_test,y_pred)
print("Confusion matrix:\n",cm)
p = cm[1][1]/(cm[1][1]+cm[0][1])
r = cm[1][1]/(cm[1][1]+cm[1][0])
print("precision: ",p)
print("recall: ",r)
print("F Score",(2*p*r)/(p+r))


