#Remove the existing Environment-
rm(list=ls())

#Set working directory-
setwd("D:/R-programming/1.Project- Employee Absenteeism")

#Check the working directory-
getwd()

#Load the Data-
library(xlsx)   #Library for read excel file
data= read.xlsx("Absenteeism_at_work_Project.xls",sheetIndex = 1,header=TRUE)

#---------------------------Exploratory Data Analysis--------------------------------------#

head(data)
class(data)
dim(data)
str(data)
names(data)
summary(data)

#from the data summary we can see variable ID, which is not useful variable 
#in modeling further, so here removing variable ID from dataset.
data= subset(data,select=-c(ID))

#since month variable can contain maximum 12 values, so here replace 0 with NA-
data$Month.of.absence[data$Month.of.absence %in% 0]= NA

#Dividing Work_load_Average/day_ by 1000 (As told by the support team)
data$Work.load.Average.day.= data$Work.load.Average.day./1000

#Extract column names of numeric and categorical variables-
cnames = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
                    'Work.load.Average.day.', 'Transportation.expense',
                    'Hit.target', 'Weight', 'Height', 
                    'Body.mass.index', 'Absenteeism.time.in.hours')

cat_cnames = c('Reason.for.absence','Month.of.absence','Day.of.the.week',
                     'Seasons','Disciplinary.failure', 'Education', 'Social.drinker',
                     'Social.smoker', 'Son', 'Pet')


#============================Data Pre-processing==================================================#

#---------------------------Missing Value Analysis------------------------------------------------#
#Total missing values in dataset-
sum(is.na(data))

#check missing values in target variable-
sum(is.na(data$Absenteeism.time.in.hours))
#remove the observations in which target varia le have missing value-
data= data[(!data$Absenteeism.time.in.hours %in% NA),]

#remaining missing values in data-
sum(is.na(data))

#Calculate missing values in dataset-
missing_value= data.frame(apply(data,2,function(x)sum(is.na(x))))
missing_value$variable= row.names(missing_value)
row.names(missing_value)=NULL
names(missing_value)[1]="missing_precentage"
missing_value= missing_value[,c(2,1)]
missing_value$missing_precentage= (missing_value$missing_precentage/nrow(data))*100
missing_value= missing_value[order(-missing_value$missing_precentage),]
write.csv(missing_value,"missing_value.csv", row.names = FALSE)

#Missing value imputation for categorical variables-
#Mode method-
mode=function(v){
  uniqv=unique(v)
  uniqv[which.max(tabulate(match(v,uniqv)))]
}

for(i in cat_cnames){
  print(i)
  data[,i][is.na(data[,i])] = mode(data[,i])
}

#Now check the remaining missing values-
sum(is.na(data))
#missing values= 72

#Missing value imputation for numeric variables-
#Lets take one sample data for referance-

data$Body.mass.index[11]

#Actual value= 23
#Mean= 26.68
#Median= 25
#KNN= 23

#Mean method-
#repalce the sample data with NA to check the accuracy to select the final method-
data$Body.mass.index[11]=NA
data$Body.mass.index[is.na(data$Body.mass.index)] = mean(data$Body.mass.index,
                                                         na.rm=TRUE)
data$Body.mass.index[11]
#Mean = 26.68

#Median Method- #reload the data first
data$Body.mass.index[11]=NA
data$Body.mass.index[is.na(data$Body.mass.index)]= median(data$Body.mass.index,na.rm=TRUE)
data$Body.mass.index[11]
#Median= 25

#KNN Imputation- #reload the data first
data$Body.mass.index[11]=NA
library(DMwR)  #Library for KNN
data= knnImputation(data,k=3)
data$Body.mass.index[11]  
#KNN=23  

#-> From all above method we have seen that KNN is more accurate then all other method,
#So we will take KNN Imputation for missing value imputation.
sum(is.na(data))

#-> Now, here data is free from missing values.

#-----------------------------Outlier Analysis-------------------------------------------------------#
#save data for reference-
df= data
data=df

#Create box-plot for outlier analysis-
library(ggplot2)    #Library for visualization
for(i in 1:length(cnames))
  assign(paste0("AB",i),ggplot(aes_string(y=(cnames[i]),x="Absenteeism.time.in.hours"),
                               d=subset(data))
  +geom_boxplot(outlier.colour = "Red",outlier.shape = 18,outlier.size = 2,
                fill="skyblue4")+theme_gray()
  +stat_boxplot(geom = "errorbar", width=0.5)
  +labs(y=cnames[i],x="Absenteeism time(Hours)")
  +ggtitle("Box Plot of Absenteeism for",cnames[i]))
  
gridExtra::grid.arrange(AB1,AB2,ncol=2)
gridExtra::grid.arrange(AB3,AB4,ncol=2)
gridExtra::grid.arrange(AB5,AB6,ncol=2)
gridExtra::grid.arrange(AB7,AB8,ncol=2)
gridExtra::grid.arrange(AB9,AB10,ncol=2)

#Remove outliers from dataset-
for(i in cnames){
  print(i)
  outlier= data[,i][data[,i] %in% boxplot.stats(data[,i])$out]
  print(length(outlier))
  data=data[which(!data[,i] %in% outlier),]
}  

#Replace outliers with NA and impute using KNN method-
for(i in cnames){
  print(i)
  outlier= data[,i][data[,i] %in% boxplot.stats(data[,i])$out]
  print(length(outlier))
  data[,i][data[,i] %in% outlier]=NA
}
sum(is.na(data))

#KNN-
data= knnImputation(data,k=3)
sum(is.na(data)) 

#check the outlier after imputation-
for(i in 1:length(cnames))
  assign(paste0("AB",i),ggplot(aes_string(y=(cnames[i]),x="Absenteeism.time.in.hours"),
                               d=subset(data))
         +geom_boxplot(outlier.colour = "Red",outlier.shape = 18,outlier.size = 2,
                       fill="skyblue4")+theme_gray()
         +stat_boxplot(geom = "errorbar", width=0.5)
         +labs(y=cnames[i],x="Absenteeism time(Hours)")
         +ggtitle("Box Plot of Absenteeism for",cnames[i]))

gridExtra::grid.arrange(AB1,AB2,ncol=2)
gridExtra::grid.arrange(AB3,AB4,ncol=2)
gridExtra::grid.arrange(AB5,AB6,ncol=2)
gridExtra::grid.arrange(AB7,AB8,ncol=2)
gridExtra::grid.arrange(AB9,AB10,ncol=2)

#-> Now, here data is free from outliers.

df1=data

#--------------------------------Feature Selection--------------------------------------------#

df=data
data=df

#Correlation Analysis for continuous variables-
library(corrgram)    #Library for correlation plot

corrgram(data[,cnames],order=FALSE,upper.panel = panel.pie,
           text.panel = panel.txt,font.labels =1,
           main="Correlation plot for Absenteeism")

#Correlated variable= weight & Body mass index.

#Anova Test for categorical variable-

for(i in cat_cnames){
  print(i)
  Anova_result= summary(aov(formula = Absenteeism.time.in.hours~data[,i],data))
  print(Anova_result)
}

#redudant categorical variables- Pet,Social.smoker,Education,Seasons,Month.of.absence

#Dimensionity Reduction
data= subset(data,select=-c(Weight,Pet,Social.smoker,Education,Seasons,Month.of.absence))
dim(data)

#---------------------------------Feature Scaling--------------------------------------------------#

df= data
data=df

#update the continuous variable-
cnames= c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
           'Work.load.Average.day.', 'Transportation.expense',
           'Hit.target','Height', 'Body.mass.index','Absenteeism.time.in.hours')

#update the categorical variable-
cat_cnames = c('Reason.for.absence','Day.of.the.week',
              'Disciplinary.failure','Social.drinker','Son')
           

#summary of data to check min and max values of numeric variables-
summary(data)

#Skewness of numeric variables-
library(propagate)

for(i in cnames){
  skew = skewness(data[,i])
  print(i)
  print(skew)
}

#log transform
data$Absenteeism.time.in.hours = log1p(data$Absenteeism.time.in.hours)

#Normality check-
hist(data$Absenteeism.time.in.hours,col="Yellow",main="Histogram of Absenteeism ")
hist(data$Distance.from.Residence.to.Work,col="Red",main="Histogram of 
     distance between work and residence")
hist(data$Transportation.expense,col="Green",main="Histogram of Transportation Expense")

#From all above histogram plot we can say that data is not uniformaly distributed,
#So best method for scaling will be normalization-

#Normalization-
for(i in cnames){
  if(i !='Absenteeism.time.in.hours'){
    print(i)
  data[,i]= (data[,i]-min(data[,i]))/(max(data[,i]-min(data[,i])))
  print(data[,i])
 }
}

#Summary of data after all preprocessing-
summary(data)

write.csv(data,"Absenteeism_Pre_processed_Data.csv",row.names=FALSE)


#===============================Model Devlopment=================================================#

#Clean the Environment-
library(DataCombine)
rmExcept("data")

#Data Copy for refrance-
df=data
data=df

cat_cnames = c('Reason.for.absence','Day.of.the.week',
               'Disciplinary.failure','Social.drinker','Son')

#create dummy variable for categorical variables-
library(dummies)
data = dummy.data.frame(data, cat_cnames)

dim(data)

#Divide the data into train and test-
set.seed(6789)
train_index= sample(1:nrow(data),0.8*nrow(data))
train= data[train_index,]
test= data[-train_index,]

#--------------------------Decision Tree for Regression------------------------------------------#

#Model devlopment for train data-
library(rpart)    #Library for regression model
DT_model= rpart(Absenteeism.time.in.hours~.,train,method="anova")
DT_model

#Prediction for train data-
DT_train=predict(DT_model,train[-51])

#Prediction for test data-
DT_test=predict(DT_model,test[-51])

#Error metrics to calculate the performance of model-
rmse= function(y,y1){
  sqrt(mean(abs(y-y1)^2))
}

#RMSE calculation for train data-
rmse(train[,51],DT_train)
#RMSE_train=  0.4303573

#RMSE calculation for test data-
rmse(test[,51],DT_test)
#RMSE_test= 0.4250704

#r-square calculation-
#function for r-square-
rsquare=function(y,y1){
  cor(y,y1)^2
}

#r-square calculation for train data-
rsquare(train[,51],DT_train)
#r-square_train= 0.5594418

#r-square calculation for test data-
rsquare(test[,51],DT_test)
#r-square_test= 0.6366973

#Visulaization to check the model performance on test data-
plot(test$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green",main="Decision Tree")
lines(DT_test,type="l",col="Blue")

#Write rule into drive-
write(capture.output(summary(DT_model)),"Decision_Tree_Model.txt")

#----------------------------Random Forest for Regression------------------------------------------#

library(randomForest)  #Library for randomforest machine learning algorithm
library(inTrees)       #Library for intree transformation
RF_model= randomForest(Absenteeism.time.in.hours~.,train,ntree=300,method="anova")

#transform ranfomforest model into treelist-
treelist= RF2List(RF_model)

#Extract rules-
rules= extractRules(treelist,train[-51])
rules[1:5,]
#covert rules into redable format-
readable_rules= presentRules(rules,colnames(train))
readable_rules[1:5,]
#Get Rule metrics-
rule_metrics= getRuleMetric(rules,train[-51],train$Absenteeism.time.in.hours)
rule_metrics= presentRules(rule_metrics,colnames(train))
rule_metrics[1:10,]
summary(rule_metrics)

#Check model performance on train data-
RF_train= predict(RF_model,train[-51])

#Check model performance on test data-
RF_test= predict(RF_model,test[-51])

#RMSE calculation for train data-
rmse(train[,51],RF_train)
#RMSE_train=  0.2386952

#RMSE calculation for test data-
rmse(test[,51],RF_test)
#RMSE_test=  0.404051

#r-square calculation for train data-
rsquare(train[,51],RF_train)
#r-square= 0.8821434

#r-square calculation for test data-
rsquare(test[,51],RF_test)
#r-square= 0.6826756

#Visulaization to check the model performance on test data-
plot(test$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green",main="Random Forest")
lines(RF_test,type="l",col="Blue")

#write rule into drive-
write(capture.output(summary(rule_metrics)),"Random_Forest_Model.txt")

#-------------------------------Linear Regression---------------------------------------------------#

#recall numeric variables to check the VIF-
numeric_index1= c("Transportation.expense","Distance.from.Residence.to.Work","Service.time",
                  "Age","Work.load.Average.day.","Hit.target","Height",
                  "Body.mass.index","Absenteeism.time.in.hours")
numeric_data1= data[,numeric_index1]
cnames1= colnames(numeric_data1)
cnames1

library(usdm)  #Library for VIF(Variance Infleation factor)
vif(numeric_data1)
vifcor(numeric_data1,th=0.7) #VIF calculation for numeric variables

#Linear regression model-
lr_model= lm(Absenteeism.time.in.hours~.,train)
summary(lr_model)

#check model performance on train data-
lr_train= predict(lr_model,train[-51])

#check model performance on test data-
lr_test= predict(lr_model,test[-51])

#RMSE calculation for train data-
rmse(train[,51],lr_train)
#RMSE_train=0.4235679

#RMSE calculation for test data-
rmse(test[,51],lr_test)
#RMSE_test=0.4375786

#r-square calculation for train data-
rsquare(train[,51],lr_train)
#r-square_train=0.5732328

#r-square calculation for test data-
rsquare(test[,51],lr_test)
#r-square_test=0.6218227

#Visulaization to check the model performance on test data-
plot(test$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green",main="Linear Regression")
lines(lr_test,type="l",col="Blue")

write(capture.output(summary(lr_model)),"Linear_Regression_Model.txt")

#--------------------------------Gradient Boosting----------------------------------------#
library(gbm)

#Develop Model
GB_model = gbm(Absenteeism.time.in.hours~., data = train, n.trees = 500, interaction.depth = 2)

#check model performance on train data-
GB_train = predict(GB_model, train,n.trees = 500)

#check model performance on test data-
GB_test = predict(GB_model, test, n.trees = 500)

#RMSE calculation for train data-
rmse(train[,51],GB_train)
#RMSE_train=0.3577656

#RMSE calculation for test data-
rmse(test[,51],GB_test)
#RMSE_test=0.4150737

#r-square calculation for train data-
rsquare(train[,51],GB_train)
#r-square_train=0.7010977

#r-square calculation for test data-
rsquare(test[,51],GB_test)
#r-square_test=0.6648008

#Visulaization to check the model performance on test data-
plot(test$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green",
     main="Gradient Boosting")
lines(GB_test,type="l",col="Blue")

#####################################Thank You##################################################

