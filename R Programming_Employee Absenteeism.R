### clear environment ###
rm(list=ls())

### load the libraries  ###

library(ggplot2)
library(corrgram)
library(corrplot)


###set working directory###

setwd("D:/DATA SCIENCE STUDY METERIAL/Projects/Employee Absenteeism_Project")
getwd()


### Load Bike renting Data CSV file ###
library(xlsx)
df=read.xlsx("DATA set.xls",sheetIndex = 1)

head(df)

###########################################################################
####################   1.3 Expletory Data Analysis   ######################
###########################################################################


dim(df)      # checking the dimension of data frame.

str(df)      # checking datatypes of all columns.

# drop the observation where Absenteeism time in hour is NAN

df= df[(!df$Absenteeism.time.in.hours %in% NA),]


#store categorical and continuous Variable column names
col=colnames(df)
print(col)  

cat_var=c('ID','Reason.for.absence','Month.of.absence','Day.of.the.week','Seasons','Disciplinary.failure',
          'Education','Son','Social.drinker','Social.smoker', 'Pet')

con_var=c('Transportation.expense','Distance.from.Residence.to.Work','Service.time','Age',
          'Work.load.Average.day.','Hit.target','Weight','Height','Body.mass.index','Absenteeism.time.in.hours')


###########################################################################
####################   2.1.	Data Preprocessing   ##########################
###########################################################################


###################  2.1.1.	Missing Value Analysis  #######################

#Calculate missing values

miss_val = data.frame(apply(df,2,function(x){sum(is.na(x))}))
miss_val
miss_val$Columns = row.names(miss_val)
row.names(miss_val)=NULL
names(miss_val)[1]="Missing_percentage"
miss_val= miss_val[,c(2,1)]
miss_val$Missing_percentage = (miss_val$Missing_percentage/nrow(df)) * 100
miss_val = miss_val[order(-miss_val$Missing_percentage),]

#Missing value imputation
data=df
data$Body.mass.index[11]                  #Lets take one sample data for referance

#Actual value= 23
#Mean= 26.71
#Median= 25
#KNN= 23

#Mean method-
data$Body.mass.index[11]=NA
data$Body.mass.index[is.na(data$Body.mass.index)] = mean(data$Body.mass.index,
                                                         na.rm=TRUE)
data$Body.mass.index[11]
#Mean = 26.71

#Median Method- 
data$Body.mass.index[11]=NA
data$Body.mass.index[is.na(data$Body.mass.index)]= median(data$Body.mass.index,na.rm=TRUE)
data$Body.mass.index[11]
#Median= 25

#KNN Imputation- #reload the data first
data$Body.mass.index[11]=NA
library(DMwR)  #Library for KNN
data= knnImputation (data,k = 5)
data$Body.mass.index[11]  
#KNN=23

###################  2.1.2. outlier analysis #############################


#create Box plot for outlier analysis
cnames= con_var
for(i in 1:length(cnames)){
  assign(paste0("AB",i),ggplot(aes_string(x="Absenteeism.time.in.hours",y=(cnames[i])),data=subset(df))+
           geom_boxplot(outlier.color = "Red",outlier.shape = 18,outlier.size = 2,
                        fill="Purple")+theme_get()+
           stat_boxplot(geom = "errorbar",width=0.5)+
           labs(x="Absenteeism.time.in.hours",y=cnames[i])+
           ggtitle("Boxplot o f",cnames[i]))
}

gridExtra::grid.arrange(AB1,AB2,AB3,AB4,AB5,ncol=5)   # plot all graph 
gridExtra::grid.arrange(AB6,AB7,AB8,AB9,AB10,ncol=5)
#Replace outliers with NA

for(i in con_var){
  print(i)
  outlier= df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  print(length(outlier))
  df[,i][df[,i] %in% outlier]=NA
}

sum(is.na(df))

#Impute outliers by KNN method

data= knnImputation (data,k = 5)


df=data

#########################  2.1.3. Data Observation  ############################

#convert data into proper data type after KNN imputation

for (i in col){
  df[,i]=as.integer(df[,i])
}

for (i in cat_var){
  df[,i]=as.factor(df[,i])
  
}

str(df)                                 #checking datatypes of all columns

summary(df[,con_var])                   # checking numerical variables
summary(df[,cat_var])                   # checking categorical variables


# final data after missing value and outlier analysis

write.csv(df,"Data after missing value and outlier.csv", row.names=FALSE )



#############################################################################
########################## VISULAIZATION ####################################
#############################################################################

# bar graph for categorical variables

plot_bar <- function(cat, y, fun){
  gp = aggregate(x = df[, y], by=list(cat=df[, cat]), FUN=fun)
  ggplot(gp, aes_string(x = 'cat', y = 'x'))+
    geom_bar(stat = 'identity',fill = "aquamarine3")+
    labs(y = y, x = cat)+theme(panel.background = element_rect("antiquewhite"))+
    theme(plot.title = element_text(size = 9))+
    ggtitle(paste("Bar plot for",y,"wrt to",cat))
}

for(i in 1:length(cat_var)) {
  assign(paste0("PB",i),plot_bar(cat_var[i],'Absenteeism.time.in.hours','sum'))
}

gridExtra::grid.arrange(PB1,PB2,PB3,ncol=3)
gridExtra::grid.arrange(PB4,PB5,PB6,ncol=3)
gridExtra::grid.arrange(PB7,PB8,PB9,ncol=3)



# histogram for continuous variables

hist_plot <- function(column, dataset){
  hist(x=dataset[,column],col="Green",xlab=column,ylab="density",
       main=paste("Histogram of ", column))
}

for(i in 1:length(con_var)) {
  assign(paste0("PH",i),hist_plot(con_var[i],dataset= df))
}


# Chacking VIF for skewness

cnames=con_var
library(propagate)
for(i in cnames){
  print(i)
  skew= skewness(df[,i])
  print(skew)
}                                      

#########################  2.1.4.Feature Selection  ###############################


####  for continuous variables ####

# correlation plot for numerical feature

corrgram(df[,con_var], order = FALSE,
         upper.panel = panel.cor, text.panel = panel.txt,
         main = "Correlation Plot")



####  for categorical Variable  ####

#Anova analysis for categorical variable with target numeric variable-
for(i in cat_var){
  print(i)
  Anova_result= summary(aov(formula = Absenteeism.time.in.hours~df[,i],df))
  print(Anova_result)
}

#########################  2.1.5. Feature Scaling  ###############################

#as the variables are not unifomaly distributed we use normalization for Feature scaling

#Normalization-

for(i in con_var){
    print(i)
    df[,i]= (df[,i]-min(df[,i]))/(max(df[,i]-min(df[,i])))
    print(df[,i])
  }



######################  2.1.5.	Data after EDA and preprocessing  ################


df= subset(df,select= -c (Weight,Day.of.the.week,Social.smoker,Education,Seasons,Pet))


head(df)
dim(df)
str(df)

write.csv(df,"Absenteeism_Pre_processed_Data.csv", row.names=FALSE )

# change categorical to numeric making bin for regression model

cat_index=sapply(df,is.factor)
cat_data=df[,cat_index]
cat_var=colnames(cat_data)
library(dummies)
df= dummy.data.frame(df,cat_var)


###########################################################################
####################   2.2. Model Development   ###########################
###########################################################################


#########################  2.2.1 Model building  #########################

#clear all the data except final data set.

data=df
df=data

library(DataCombine)
rmExcept("data")


#Function for Error metrics to calculate the performance of model-

rmse= function(y,y1){
  sqrt(mean(abs(y-y1)^2))
}

#Function for r2 to calculate the goodness of fit of model-
rsquare=function(y,y1){
  cor(y,y1)^2
}


# devide the data in train and test

set.seed(123)
train_index= sample(1:nrow(data),0.8*nrow(data))
train= data[train_index,]
test= data[-train_index,]


################## 2.2.1.desision tree for regression  ##########################

library(rpart)

fit=rpart(Absenteeism.time.in.hours~.,data=train,method = "anova") #model development on train data

DT_test=predict(fit,test[,-96])            #predict test data
DT_train= predict(fit,train[,-96])         #predict train data

DT_RMSE_Test = rmse(test[,96],DT_test)     # RMSE calculation for test data
DT_RMSE_Train = rmse(train[,96],DT_train)  # RMSE calculation for train data

DT_r2_test=rsquare(test[,96],DT_test)      # r2 calculation for test data
DT_r2_train= rsquare(train[,96],DT_train)  # r2 calculation for train data


### 2.2.2. Random forest for regression ###

library(randomForest)

RF_model= randomForest(Absenteeism.time.in.hours~.,train,ntree=100,method="anova") #Model development on train data

RF_test= predict(RF_model,test[-96])      #Prediction on test data
RF_train= predict(RF_model,train[-96])    #Prediction on train data

RF_RMSE_Test=rmse(test[,96],RF_test)      #RMSE calculation of test data-
RF_RMSE_Train=rmse(train[,96],RF_train)   #RMSE calculation of train data

RF_r2_test=rsquare(test[,96],RF_test)     #r2 calculation for test data-
RF_r2_train=rsquare(train[,96],RF_train)  #r2 calculation for train data-




###  2.2.3. Linear Regression ###


LR_model= lm(Absenteeism.time.in.hours~.,train)               #Model devlopment on train data
summary(LR_model)

LR_test= predict(LR_model,test[-96])      #prediction on test data
LR_train= predict(LR_model,train[-96])    #prediction on train data

LR_RMSE_Test=rmse(test[,96],LR_test)      #RMSE calculation of test data
LR_RMSE_Train=rmse(train[,96],LR_train)   #RMSE calculation of train data

LR_r2_test=rsquare(test[,96],LR_test)     #r2 calculation for test data
LR_r2_train=rsquare(train[,96],LR_train)  #r2 calculation for train data



### 2.2.4. Gradient Boosting ###

library(gbm)

GB_model = gbm(Absenteeism.time.in.hours~., data = train, n.trees = 100, interaction.depth = 2) #Model devlopment on train data

GB_test = predict(GB_model, test[-96], n.trees = 100)      #prediction on test data
GB_train = predict(GB_model, train[-96], n.trees = 100)    #prediction on train data

GB_RMSE_Test=rmse(test[,96],GB_test)                       
GB_RMSE_Train=rmse(train[,96],GB_train)                    #Mape calculation of train data

GB_r2_test=rsquare(test[,96],GB_test)                      #r2 calculation for test data-
GB_r2_train=rsquare(train[,96],GB_train)                   #r2 calculation for train data-




Result= data.frame('Model'=c('Decision Tree for Regression','Random Forest',
                             'Linear Regression','Gradient Boosting'),
                   'RMSE_Train'=c(DT_RMSE_Train,RF_RMSE_Train,LR_RMSE_Train,GB_RMSE_Train),
                   'RMSE_Test'=c(DT_RMSE_Test,RF_RMSE_Test,LR_RMSE_Test,GB_RMSE_Test),
                   'R-Squared_Train'=c(DT_r2_train,RF_r2_train,LR_r2_train,GB_r2_train),
                   'R-Squared_Test'=c(DT_r2_test,RF_r2_test,LR_r2_test,GB_r2_test))

Result           #Random forest and Gradient Bosting have best fit model for the data.


#########################  2.2.2.	Hyperparameter Tuning  #########################

#Random Search CV in Random Forest

library(caret)

control = trainControl(method="repeatedcv", number=3, repeats=1,search='random')


RRF_model = caret::train(Absenteeism.time.in.hours~., data=train, method="rf",trControl=control,tuneLength=1)   #model devlopment on train data
best_parameter = RRF_model$bestTune                            #Best fit parameters
print(best_parameter)
#mtry=17 As per the result of best_parameter

RRF_model = randomForest(Absenteeism.time.in.hours ~ .,train, method = "rf", mtry=17,importance=TRUE)        #build model based on best fit

RRF_test= predict(RRF_model,test[-96])                        #Prediction on test data
RRF_train= predict(RRF_model,train[-96])                      #Prediction on train data


RRF_RMSE_Test = rmse(test[,96],RRF_test)                      #Mape calculation of test data
RRF_RMSE_Train = rmse(train[,96],RRF_train)                   #Mape calculation of train data


RRF_r2_test=rsquare(test[,96],RRF_test)                      #r2 calculation for test data
RRF_r2_train= rsquare(train[,96],RRF_train)                  #r2 calculation for train data


# Grid Search CV in Random Forest

control = trainControl(method="repeatedcv", number=3, repeats=3, search="grid")
tunegrid = expand.grid(.mtry=c(6:18))


GRF_model= caret::train(Absenteeism.time.in.hours~.,train, method="rf", tuneGrid=tunegrid, trControl=control)     #model devlopment on train data
best_parameter = GRF_model$bestTune                           #Best fit parameters
print(best_parameter)
#mtry=8 As per the result of best_parameter

GRF_model = randomForest(Absenteeism.time.in.hours ~ .,train, method = "anova", mtry=9)                    #build model based on best fit

GRF_test= predict(GRF_model,test[-96])                      #Prediction on test data
GRF_train= predict(GRF_model,train[-96])                    #Prediction on train data


GRF_RMSE_Test = rmse(test[,96],GRF_test)                    #Mape calculation of test data
GRF_RMSE_Train = rmse(train[,96],GRF_train)                 #Mape calculation of train data


GRF_r2_test=rsquare(test[,96],GRF_test)                     #r2 calculation for test data
GRF_r2_train= rsquare(train[,96],GRF_train)                 #r2 calculation for train data



final_result= data.frame('Model'=c('Decision Tree for Regression','Random Forest','Linear Regression',
                                   'Gradient Boosting','Random Search CV in Random Forest',
                                   'Grid Search CV in Random Forest'),
                         'RMSE_Train'=c(DT_RMSE_Train,RF_RMSE_Train,LR_RMSE_Train,GB_RMSE_Train,
                                        RRF_RMSE_Train,GRF_RMSE_Train),
                         'RMSE_Test'=c(DT_RMSE_Test,RF_RMSE_Test,LR_RMSE_Test,GB_RMSE_Test,
                                       RRF_RMSE_Test,GRF_RMSE_Test),
                         'R-Squared_Train'=c(DT_r2_train,RF_r2_train,LR_r2_train,GB_r2_train,
                                             RRF_r2_train,GRF_r2_train),
                         'R-Squared_Test'=c(DT_r2_test,RF_r2_test,LR_r2_test,GB_r2_test,
                                            RRF_r2_test,GRF_r2_test))

print(final_result)

###################################################################################################
#2. How much losses every month can we project in 2011 if same trend of absenteeism continues?
###################################################################################################
df1= read.csv('Data after missing value and outlier.csv')
colnames(df1)
data_work_loss=subset(df1,select= c(Month.of.absence, Work.load.Average.day., Absenteeism.time.in.hours))
data_work_loss$Work_loss_per_day=
  (data_work_loss$Work.load.Average.day./24)*data_work_loss$Absenteeism.time.in.hours
Monthly_loss= aggregate(data_work_loss$Work_loss_per_day, by= list(data_work_loss$Month.of.absence), FUN=sum)
Monthly_loss
