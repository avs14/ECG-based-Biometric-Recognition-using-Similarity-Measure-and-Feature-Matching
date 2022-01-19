import pandas as pd
from sklearn.tree import DecisionTreeClassifier # Import Decision Tree Classifier
from sklearn.model_selection import train_test_split # Import train_test_split function
from sklearn import metrics

import numpy as np

col = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N']

pima = pd.read_csv(r'data2/testing2.csv',header=None, names=col)
pima.head()

features = ['C','D','E','F','H','I','J','K','M']

X = pima[features]
y = pima.N

print(X)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=3)

clf = DecisionTreeClassifier(criterion="entropy", max_depth=3)
clf = clf.fit(X_train.values,y_train.values)
y_pred = clf.predict(X_test.values)
print(len(y_test))

print(len(y_pred))

print("Accuracy:",metrics.accuracy_score(y_test, y_pred))
cm = metrics.confusion_matrix(y_test,y_pred)
print("Confusion matrix:\n",cm)
p = cm[1][1]/(cm[1][1]+cm[0][1])
r = cm[1][1]/(cm[1][1]+cm[1][0])
print("precision: ",p)
print("recall: ",r)
print("F Score",(2*p*r)/(p+r))


test = np.array([87.62645506108203,9.36090033389321,0.99915168969483594, 334.25841760615685,48.246046845953245,0.03675921259181475,0.030431365083580384,0.9642904903353721,0.35]).reshape(1,-1)
y_pred = clf.predict(test)
print("\n")
print(y_pred)