#Clear the environment
rm(list = ls())

#Set working Directory
setwd("/Users/raghavkotwal/Documents/Data Science/Employee Absenteeism")
getwd()

#Load the librarires
libraries = c("dummies","caret","rpart.plot","plyr","dplyr", "ggplot2","rpart","dplyr","DMwR","randomForest","usdm","corrgram","DataCombine","xlsx","tree")
lapply(X = libraries,require, character.only = TRUE)
rm(libraries)

#Read the Data
df = read.xlsx(file = "Absenteeism_at_work.xls", header = T, sheetIndex = 1)


######Explore Data#######

#Check Dimensions of Data
dim(df)

#Check Structure of Variables
str(df)

#view the top 5 rows of the data
head(df)

#Transform required data types into categorical data
df$Reason.for.absence[df$Reason.for.absence %in% 0] = NA
df$Reason.for.absence = as.factor(as.character(df$Reason.for.absence))

df$Month.of.absence[df$Month.of.absence %in% 0] = NA
df$Month.of.absence = as.factor(as.character(df$Month.of.absence))

df$Day.of.the.week = as.factor(as.character(df$Day.of.the.week))
df$Seasons = as.factor(as.character(df$Seasons))
df$Disciplinary.failure = as.factor(as.character(df$Disciplinary.failure))
df$Education = as.factor(as.character(df$Education))
df$Son = as.factor(as.character(df$Son))
df$Social.drinker = as.factor(as.character(df$Social.drinker))
df$Social.smoker = as.factor(as.character(df$Social.smoker))
df$Pet = as.factor(as.character(df$Pet))




##########MISSING VALUE ANALYSIS#############

sapply(df,function(x){sum(is.na(x))})
missing_values = data.frame(sapply(df,function(x){sum(is.na(x))}))

#Calculate missing percentage and arrage in order
missing_values$Var = row.names(missing_values)
row.names(missing_values) = NULL
names(missing_values)[1] = "Percentage"
missing_values$Percentage = ((missing_values$Percentage/nrow(df)) *100)
missing_values = missing_values[,c(2,1)]
missing_values = missing_values[order(-missing_values$Percentage),]

#Create missing value and impute using mean, median and knn
#Value = 31
#Mean = 26.68
#Median = 25
#KNN = 31
#df1[["Body.mass.index"]][6]
#df1[["Body.mass.index"]][6] = NA
#df1[["Body.mass.index"]][6] = mean(df$Body.mass.index, na.rm = T)
#df1[["Body.mass.index"]][6] = median(df$Body.mass.index, na.rm = T)

df = knnImputation(data = df, k = 5)

#Check if any missing values
sum(is.na(df))




#Check Structure of Variables
str(df)
df1 = df



####### Data visualisation #####

library(ggplot2)
library(corrplot)

#Definig variable types
numerical_set = c("ID","Transportation.expense","Distance.from.Residence.to.Work","Service.time","Age","Work.load.Average.day.","Hit.target","Son","Pet","Height","Weight","Body.mass.index","Absenteeism.time.in.hours")
categorical_set = c("Reason.for.absence","Month.of.absence","Day.of.the.week","Seasons","Disciplinary.failure","Education","Social.drinker","Social.smoker")


## plotting numerical data set vs the target variable
plot(df$ID,df$Absenteeism.time.in.hours,x = "ID",ylab = "Absenteeism time in hours",main = "Absenteeism time vs ID",col="Black")
plot(df$Transportation.expense,df$Absenteeism.time.in.hours,xlab = "Transportation expense",ylab = "Absenteeism time in hours",main = "Absent time vs Transp expense",col="Black")
plot(df$Distance.from.Residence.to.Work,df$Absenteeism.time.in.hours,xlab = "Distance.from.Residence.to.Work",ylab = "Absenteeism time in hours",main = "Absent time vs Travel distance",col="Black")
plot(df$Service.time,df$Absenteeism.time.in.hours,xlab = "Service Time",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Service time",col="Black")
plot(df$Age,df$Absenteeism.time.in.hours,xlab = "Age",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Age",col="Black")
plot(df$Work.load.Average.day.,df$Absenteeism.time.in.hours,xlab = "Work.load.Average.day",ylab = "Absenteeism time in hours",main = "Absent time vs Work.load.Avg.day",col="Black")
plot(df$Hit.target,df$Absenteeism.time.in.hours,xlab = "Hit target",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Hit target",col="Black")
plot(df$Son,df$Absenteeism.time.in.hours,xlab = "Son",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Son",col="Black")
plot(df$Pet,df$Absenteeism.time.in.hours,xlab = "Pet",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Pet",col="Black")
plot(df$Height,df$Absenteeism.time.in.hours,xlab = "Height",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Height",col="Black")
plot(df$Weight,df$Absenteeism.time.in.hours,xlab = "Weight",ylab = "Absenteeism time in hours",main = "Absenteeism time vs Weight",col="Black")
plot(df$Body.mass.index,df$Absenteeism.time.in.hours,xlab = "BMI",ylab = "Absenteeism time in hours",main = "Absenteeism time vs BMI",col="Black")



##plotting categorical data set vs target variable
dev.off()
for(i in 1:length(categorical_set)){
  assign(paste0("gg",i),ggplot(aes_string(y=df$Absenteeism.time.in.hours,x=df[,categorical_set[i]]),data = subset(df))
         + stat_boxplot(geom = "errorbar",width = 0.3) +
           geom_boxplot(outlier.colour = "red",fill = "black",outlier.shape = 18,outlier.size = 1) +
           labs(y = "Absenteeism.time.in.hours",x=names(df[categorical_set[i]])) + 
           ggtitle(names(df[categorical_set[i]])))
  
}
gridExtra::grid.arrange(gg1,gg2,nrow = 2,ncol=1)
gridExtra::grid.arrange(gg3,gg4,nrow = 2,ncol = 1)
gridExtra::grid.arrange(gg5,gg6,nrow = 2,ncol = 1)
gridExtra::grid.arrange(gg7,gg8,nrow = 2,ncol = 1)





####Outlier Analysis#####

## Replace outliers in numerical dataset with NAs using boxplot method 
for(i in numerical_set){
  outlier_value = boxplot.stats(df[,i])$out
  print(names(df[i]))
  print(outlier_value)
  df[which(df[,i] %in% outlier_value),i] = NA
}

#Compute the NA values using KNN imputation
df = knnImputation(df, k = 5)

#Check if any missing values
sum(is.na(df))


######Feature Selection######

## Correlation Plot 
corrgram(df[,numerical_set], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

## ANOVA test for Categprical variable
summary(aov(formula = Absenteeism.time.in.hours~ID,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Reason.for.absence,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Month.of.absence,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Day.of.the.week,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Seasons,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Disciplinary.failure,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Education,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Social.drinker,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Social.smoker,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Son,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Pet,data = df))


## Dimension Reduction
df = subset(df, select = -(which(names(df) %in% c("Weight","Day.of.the.week","Seasons","Education","Social.smoker","Social.drinker"))))

## Using createdataPartition for sampling we create 75% train 25% test data set using variable reason for Absence
train = createDataPartition(df$Reason.for.absence,times = 1,p = 0.75,list = F)
test = -(train)

#####Model Development#####
## Linear regression 
modelLR = lm(Absenteeism.time.in.hours~.,data = df[train,])
summary(modelLR)
par(mar = c(2,2,2,2))
par(mfrow = c(3,1))
plot(modelLR)
predictLR = predict(modelLR,df[test,])
sqrt(mean((predictLR-df$Absenteeism.time.in.hours[test])^2))
##RMSE : 2.94




## Decision trees 
modelDT = tree(Absenteeism.time.in.hours~.,df,subset = train)
summary(modelDT)
plot(modelDT)
text(modelDT,pretty = 0)
predictDT = predict(modelDT,newdata = df[test,])
sqrt(mean((predictDT-df$Absenteeism.time.in.hours[test])^2))
##RMSE 3.06


## Random Forest
modelRF = randomForest(Absenteeism.time.in.hours~.,data = df,subset = test,mtry = 12,ntree=12,importance = TRUE)
varImpPlot(modelRF)
importance(modelRF)
predictRF = predict(modelRF,newdata = df[test,])
sqrt(mean((predictRF-df$Absenteeism.time.in.hours[test])^2))
##RMSE 1.44


#### Conclusion
## Sum and mean of the absenteeism hours reason wise
reason_sum_hrs = aggregate(df$Absenteeism.time.in.hours,by = list(Category = df$Reason.for.absence),FUN = sum)
names(reason_sum_hrs)=c("Reason no.","Sum of absent hours")
reason_mean_hrs = aggregate(df$Absenteeism.time.in.hours,by = list(Category = df$Reason.for.absence),FUN = mean)
names(reason_mean_hrs)=c("Reason no.","Mean of absent hours")
table(df$Reason.for.absence)

## Monthly loss for the Company

loss_data = df[,c("Month.of.absence","Work.load.Average.day.","Service.time","Absenteeism.time.in.hours")]
str(loss_data)
loss_data$WorkLoss = round((loss_data$Work.load.Average.day./loss_data$Service.time)*loss_data$Absenteeism.time.in.hours)
View(loss_data)
monthly_loss = aggregate(loss_data$WorkLoss,by = list(Category = loss_data$Month.of.absence),FUN = sum)
names(monthly_loss) = c("Month","WorkLoss")

