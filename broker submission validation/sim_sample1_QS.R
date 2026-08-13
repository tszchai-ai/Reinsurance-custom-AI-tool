# ============================================
# Fake Broker Submission Generator
# ============================================
library(lubridate)
library(openxlsx)
library(dplyr)
library(VGAM)

set.seed(20251018)  # reproducible, but all figures are random draws

# ---------------------------
# Global Parameters / Dates
# ---------------------------
LOSS_VAL_DATE  <- as.Date("2024-11-30")
EXPO_ASOF_DATE <- LOSS_VAL_DATE
TREATY_START   <- as.Date("2025-04-01")
TREATY_END     <- as.Date("2026-03-31")
AYs            <- 2018:2024
AYs_char       <- as.character(AYs)

BROKER_NAME  <- "Atlas Re Broking"
INSURER_NAME <- "CARE Specialty Insurance Company"
PROGRAM_NAME <- "Energy QS"
EFFECTIVE_DATE <- as.character(TREATY_START)
SUMMARY_TEXT <- "Quote Requested by 07 March 2025 for Energy QS"
LAYER_STATUS <- "Quote Requested"
# Reinsurer Quote response
QUOTE_DATE   <- ""
# Information
CURRENCY     <- "USD"
COVERAGE_EFFECTIVE_DATE <- TREATY_START
COVERAGE_EXPIRY_DATE <- TREATY_END
TERRITORIAL_SCOPE    <- "USA (emphasis CA, TX, FL, NY, GA)"
TREATY_BASIS <- "Risks Attaching basis" # or "Losses Occurring basis"
CLASS_OF_BUSINESS <- "Marine & Energy – Property"
LINE_OF_BUSINESS <- "Energy"
CONTRACT_NATURE <- "Treaty: Proportional Quota Share"
COVERAGE_TYPE <- "Property Damage and Business Interruption (PD/BI): Energy"
# Basis of Lines
BASIS_OF_LINES <- "Percentage of Order" # "Percentage of Whole" or "Part of Whole"
BASIS_OF_LINES_POT <- "Percentage of Order"
# Coverage Details 
COVERAGE_TERMS_CURRENCY <- "USD"
LIMIT_DEFINITION <- "Per Risk Cession Limit USD 25M any one risk; Annual Aggregate USD 200M"
CESSIONS_LIMIT_PERCENTAGE_BASIS <- "90% of Original Policy Limit (sum insured) per risk"
# Premium
EST_GROSS_CEDED_PREMIUM <- 1e07
# Ceding Commission
COMMISSION_TYPE <- "Flat" #"Sliding Scale", "Profit Commission"
COMMISSION_BASIS <- "On Ceded Written Premium, net of brokerage and taxes"
# Brokerage
BROKERAGE_PERCENT <- 2.5
GROSS_OR_NET <- "Gross"

# ---------------------------
# Cover / Terms / README / Dictionary
# ---------------------------
cover_info_df <- t(data.frame(
  BROKER_NAME = BROKER_NAME,
  INSURER_NAME = INSURER_NAME,
  PROGRAM_NAME = PROGRAM_NAME,
  EFFECTIVE_DATE = EFFECTIVE_DATE,
  SUMMARY_TEXT = SUMMARY_TEXT,
  LAYER_STATUS = LAYER_STATUS,
  QUOTE_DATE = QUOTE_DATE,
  CURRENCY = CURRENCY,
  COVERAGE_EFFECTIVE_DATE = COVERAGE_EFFECTIVE_DATE,
  COVERAGE_EXPIRY_DATE = COVERAGE_EXPIRY_DATE,
  TERRITORIAL_SCOPE = TERRITORIAL_SCOPE,
  TREATY_BASIS = TREATY_BASIS,
  CLASS_OF_BUSINESS = CLASS_OF_BUSINESS,
  LINE_OF_BUSINESS = LINE_OF_BUSINESS,
  CONTRACT_NATURE = CONTRACT_NATURE,
  COVERAGE_TYPE = COVERAGE_TYPE,
  BASIS_OF_LINES = BASIS_OF_LINES,
  BASIS_OF_LINES_POT = BASIS_OF_LINES_POT,
  COVERAGE_TERMS_CURRENCY = COVERAGE_TERMS_CURRENCY,
  LIMIT_DEFINITION = LIMIT_DEFINITION,
  CESSIONS_LIMIT_PERCENTAGE_BASIS = CESSIONS_LIMIT_PERCENTAGE_BASIS,
  EST_GROSS_CEDED_PREMIUM = EST_GROSS_CEDED_PREMIUM,
  COMMISSION_TYPE = COMMISSION_TYPE,
  COMMISSION_BASIS = COMMISSION_BASIS,
  BROKERAGE_PERCENT = BROKERAGE_PERCENT,
  GROSS_OR_NET = GROSS_OR_NET,
  stringsAsFactors = FALSE
))

##### Parameters for model input
## Poisson means for policy count and claim count
# lambda for policy
lambda_pol <- c(3200, 3600, 3600, 4000, 4000, 4400, 4800)
n_pol <- rpois(length(AYs_char),lambda=lambda_pol)
names(n_pol)=AYs_char
goforward_n <- rpois(1,lambda=5200)
# lambda for avg claim per policy
lambda_claim <- c(0.06, 0.06, 0.05, 0.055, 0.07, 0.07, 0.07)
n_claim <- rpois(length(AYs_char),lambda=(lambda_pol*lambda_claim))
names(n_claim)=names(lambda_claim)=AYs_char
# renewal probability
prob_renew <- c(0.00, 0.75, 0.75, 0.75, 0.75, 0.75, 0.75)
n_renew <- c(0,rbinom(length(AYs_char)-1,size=n_pol[-length(AYs_char)],prob=prob_renew[-1]))
names(n_renew)=names(prob_renew)=AYs_char
# states
states   <- c("CA","TX","FL","NY","GA",
              "IL","PA","OH","MI","NJ",
              "NC","VA","WA","AZ","MA")
states_prob <- c(5,5,3,3,2,
                 1,1,1,1,1,
                 0.5,0.5,0.5,0.5,0.5)
states_prob <- states_prob/sum(states_prob)
# Limits 
limit_options <- c(5e5, 1e6, 1.5e6, 2e6, 3e6, 5e6, 1e7)
limit_probs   <- c(0.15, 0.25, 0.45, 0.10, 0.03, 0.01, 0.01)
# GC/LS
gcls_prob <- c(0.6, 0.4)
# Typical SIC codes in Energy QS (quota share) program
energy_sic <- tibble::tribble(
  ~segment,       ~sic_code, ~sic_name,
  "Upstream",     1311, "Crude Petroleum & Natural Gas",
  "Upstream",     1321, "Natural Gas Liquids",
  "Upstream",     1381, "Drilling Oil & Gas Wells",
  "Upstream",     1382, "Oil & Gas Field Exploration Services",
  "Upstream",     1389, "Oil & Gas Field Services, NEC",
  "Midstream",    4612, "Crude Petroleum Pipelines",
  "Midstream",    4613, "Refined Petroleum Pipelines",
  "Midstream",    4922, "Natural Gas Transmission",
  "Midstream",    4923, "Natural Gas Transmission & Distribution",
  "Midstream",    4924, "Natural Gas Distribution",
  "Midstream",    4925, "Mixed/Manufactured/LP Gas Prod/Dist",
  "Downstream",   2911, "Petroleum Refining",
  "Downstream",   2992, "Lubricating Oils & Greases",
  "Downstream",   2999, "Petroleum & Coal Products, NEC",
  "Utilities",    4911, "Electric Services",
  "Utilities",    4931, "Electric & Other Services Combined",
  "Utilities",    4932, "Gas & Other Services Combined",
  "Utilities",    4939, "Combination Utilities, NEC",
  "Utilities",    4961, "Steam & Air-Conditioning Supply",
  "Wholesale",    5171, "Petroleum Bulk Stations & Terminals"
)
# Triangle dev ages
DEV_AGES <- month(LOSS_VAL_DATE) + 12*(seq(length(AYs_char)-1)-1)
## Rate monitor & premium parameters (illustrative)
rate_changes <- round(runif(length(AYs_char),-3,10),1)
names(rate_changes)=AYs_char
rate_changes
# premium factor and on-level factor (OLF)
prem_factor=c(1,cumprod(1+rate_changes[-1]/100))
names(prem_factor)=AYs_char
# Pareto quantile generator for severity
sample_severity <- function(n) rlomax(n,scale=1,shape3.q=2)
# Report-lag and settle-lag sampler (days): short/mid/long/tail mixture
sample_report_lags <- function(n) {
  mix <- sample.int(4, n, replace = TRUE, prob = c(0.7, 0.2, 0.05, 0.05))
  log.mu=log(c(45,120,240,600))[mix]
  log.sigma=c(0.35,0.35,0.4,0.5)[mix]
  ceiling(rlnorm(n, log.mu, log.sigma))}
sample_settle_lags <- function(n) {
  mix <- sample.int(4, n, replace = TRUE, prob = c(0.4, 0.3, 0.2, 0.1))
  log.mu=log(c(80,150,300,600))[mix]
  log.sigma=c(0.35,0.35,0.4,0.5)[mix]
  ceiling(rlnorm(n, log.mu, log.sigma))}
# Written premium factor adjustment
factor_GC_wp=c(1,1,1,1.0011,1.0049,1.016,1)
factor_LS_wp=c(1,1,1,1.0008,1.0037,1.012,1)

# ---------------------------
# Utility Functions
# ---------------------------
rand_date_between <- function(start, end, n) {
  start + sample.int(as.integer(end - start) + 1L, n, replace = TRUE) - 1L
}

# ---------------------------
# Build Exposure information: Output df_full.
# ---------------------------
df.list=df_new.list=df_renew.list=list()
for (ay in AYs) {
  n_pol_new_ay=max(n_pol[as.character(ay)]-n_renew[as.character(ay)],10)
  df.eff.date=rand_date_between(make_date(ay, 1, 1), make_date(ay, 12, 31), n_pol_new_ay)
  df.state=sample(states, n_pol_new_ay, replace = TRUE, prob=states_prob)
  df.gcls=sample(c("GC","LS"),n_pol_new_ay,replace = TRUE, prob=gcls_prob)
  df.sic=sample(energy_sic$sic_code,n_pol_new_ay,replace = TRUE)
  df.limit=sample(limit_options, n_pol_new_ay, replace = TRUE, prob = limit_probs)
  df.status="New"
  df.loading=prem_factor[as.character(ay)]*(1+0.25*(df.state%in%c("CA","NY","IL","AZ")))*(1+0.25*(df.gcls=="GC"))*runif(n_pol_new_ay,min=0.9,max=1.1)
  df.written.prem=round(0.004*df.loading*df.limit,2)
  df_new.list[[as.character(ay)]] <- data.frame(
    Policy_ID = "temp", #sprintf(paste0("P-",ay,"-%05d"), 1:n_pol_new_ay),
    Insured_ID = "temp", #sprintf(paste0("I-",ay,"-%05d"), 1:n_pol_new_ay),
    Policy_Year = ay,
    Treaty_Year = ay - 1*(df.eff.date<make_date(ay, month(TREATY_START), day(TREATY_START))),
    Effective_Date  = df.eff.date,
    Expiration_Date = df.eff.date%m+%years(1)-1,
    New_Renewed = df.status,
    State     = df.state,
    GC_LS     = df.gcls,
    SIC_Code  = df.sic,
    Written_Premium = df.written.prem,
    Limit     = df.limit,
    stringsAsFactors = FALSE)
  df_new.list[[as.character(ay)]]=df_new.list[[as.character(ay)]][order(df_new.list[[as.character(ay)]]$Effective_Date),]
  df_new.list[[as.character(ay)]]$Insured_ID=sprintf(paste0("I-",ay,"-%05d"), 1:n_pol_new_ay)
  if(ay>min(AYs)){
    n_pol_renew_ay=min(n_renew[as.character(ay)],n_pol[as.character(ay-1)]-10)
    df_renew.list[[as.character(ay)]]=df.list[[as.character(ay-1)]][sample(nrow(df.list[[as.character(ay-1)]]),n_pol_renew_ay,replace=F),]
    df_renew.list[[as.character(ay)]]$Effective_Date=df_renew.list[[as.character(ay)]]$Effective_Date%m+%years(1)
    df_renew.list[[as.character(ay)]]$Expiration_Date=df_renew.list[[as.character(ay)]]$Expiration_Date%m+%years(1)
    df_renew.list[[as.character(ay)]]$Policy_Year=df_renew.list[[as.character(ay)]]$Policy_Year+1
    df_renew.list[[as.character(ay)]]$Treaty_Year=df_renew.list[[as.character(ay)]]$Treaty_Year+1
    df_renew.list[[as.character(ay)]]$New_Renewed="Renewed"
    df.loading=prem_factor[as.character(ay)]*(1+0.25*(df_renew.list[[as.character(ay)]]$State%in%c("CA","NY","IL","AZ")))*(1+0.25*(df_renew.list[[as.character(ay)]]$GC_LS=="GC"))*runif(n_pol_renew_ay,min=0.9,max=1.1)
    df_renew.list[[as.character(ay)]]$Written_Premium=round(0.004*df.loading*df_renew.list[[as.character(ay)]]$Limit,2)
    df.list[[as.character(ay)]]=rbind(df_new.list[[as.character(ay)]],df_renew.list[[as.character(ay)]])
    df.list[[as.character(ay)]]=df.list[[as.character(ay)]][order(df.list[[as.character(ay)]]$Effective_Date),]
    df.list[[as.character(ay)]]$Policy_ID=sprintf(paste0("P-",ay,"-%05d"), 1:(n_pol_new_ay+n_pol_renew_ay))
    df_full=rbind(df_full,df.list[[as.character(ay)]])
  } else {
    df.list[[as.character(ay)]]=df_new.list[[as.character(ay)]]
    df.list[[as.character(ay)]]$Policy_ID=sprintf(paste0("P-",ay,"-%05d"), 1:n_pol_new_ay)
    df_full=df.list[[as.character(ay)]]
  }
}
df_full=df_full[df_full$Effective_Date<=LOSS_VAL_DATE,]
rownames(df_full)=NULL

df_policy_IF_new=df_full[(df_full$Effective_Date<=LOSS_VAL_DATE)&(df_full$Effective_Date>(LOSS_VAL_DATE-years(1))),]
df_policy_IF_old=df_full[(df_full$Effective_Date<=(LOSS_VAL_DATE-years(1)))&(df_full$Effective_Date>(LOSS_VAL_DATE-years(2))),]

# ---------------------------
# Earned Premium Simulation
# ---------------------------
EP_sim <- function(df_policy, AYs_eval, type) {
  EP_GC.vec=EP_LS.vec=numeric(length(AYs_eval))
  names(EP_GC.vec)=names(EP_LS.vec)=as.character(AYs_eval)
  for (ay in AYs_eval){
    if (type=="CY"){
      df_sel_new=df_policy[df_policy$Policy_Year==ay,]
      df_sel_old=df_policy[df_policy$Policy_Year==(ay-1),]
      md.start=as.Date("2024-01-01")
      md.end=as.Date("2024-12-31")
      if (ay==tail(AYs_eval,1)){
        ratio_new=as.numeric(make_date(ay, month(LOSS_VAL_DATE), day(LOSS_VAL_DATE)) - df_sel_new$Effective_Date + 1)/as.numeric(df_sel_new$Expiration_Date - df_sel_new$Effective_Date + 1)
        ratio_old=as.numeric(pmin(df_sel_old$Expiration_Date, make_date(ay, month(LOSS_VAL_DATE), day(LOSS_VAL_DATE))) - make_date(ay, month(md.start), day(md.start)) + 1)/as.numeric(df_sel_old$Expiration_Date - df_sel_old$Effective_Date + 1)
      } else {
        ratio_new=as.numeric(make_date(ay, month(md.end), day(md.end)) - df_sel_new$Effective_Date + 1)/as.numeric(df_sel_new$Expiration_Date - df_sel_new$Effective_Date + 1)
        ratio_old=as.numeric(df_sel_old$Expiration_Date - make_date(ay, month(md.start), day(md.start)) + 1)/as.numeric(df_sel_old$Expiration_Date - df_sel_old$Effective_Date + 1)
      }
    } else {
      df_sel_new=df_policy[df_policy$Treaty_Year==ay,]
      df_sel_old=df_policy[df_policy$Treaty_Year==(ay-1),]
      md.start=TREATY_START
      md.end=TREATY_END
      ratio_new=1
      ratio_old=0
    }
    EP_GC_new=sum(ratio_new*df_sel_new$Written_Premium*(df_sel_new$GC_LS=="GC"))
    EP_GC_old=sum(ratio_old*df_sel_old$Written_Premium*(df_sel_old$GC_LS=="GC"))
    EP_LS_new=sum(ratio_new*df_sel_new$Written_Premium*(df_sel_new$GC_LS=="LS"))
    EP_LS_old=sum(ratio_old*df_sel_old$Written_Premium*(df_sel_old$GC_LS=="LS"))
    EP_GC.vec[as.character(ay)]=EP_GC_new+EP_GC_old
    EP_LS.vec[as.character(ay)]=EP_LS_new+EP_LS_old
  }
  list(EP_GC.vec=EP_GC.vec,EP_LS.vec=EP_LS.vec)
}
# By calander year
EP_CY_new.list=EP_sim(df_policy=df_full, AYs_eval=AYs, "CY")
EP_CY_old.list=EP_sim(df_policy=df_full[df_full$Effective_Date<=(LOSS_VAL_DATE-years(1)),], AYs_eval=AYs[-length(AYs)], "CY")
# By treaty year
EP_TY_new.list=EP_sim(df_policy=df_full, AYs_eval=AYs, "TY")
EP_TY_old.list=EP_sim(df_policy=df_full[df_full$Effective_Date<=(LOSS_VAL_DATE-years(1)),], AYs_eval=AYs[-length(AYs)], "TY")
EP_TY_old.list$EP_GC.vec+EP_TY_old.list$EP_LS.vec
## output with premium factor adjustment
# old (CY)
EP_CY_old.list$EP_GC.vec
EP_CY_old.list$EP_LS.vec
# old (TY)
EP_TY_old.list$EP_GC.vec+EP_TY_old.list$EP_LS.vec
# old (TY original most recent year)
round(tail(EP_TY_old.list$EP_GC.vec+EP_TY_old.list$EP_LS.vec,1)*12/(month(LOSS_VAL_DATE)-month(TREATY_START)+1)*1.03,-5)
# old (TY revised most recent year)
round(tail(EP_TY_old.list$EP_GC.vec+EP_TY_old.list$EP_LS.vec,1)*12/(month(LOSS_VAL_DATE)-month(TREATY_START)+1),-5)
# old (TY project next year)
round(tail(EP_TY_old.list$EP_GC.vec+EP_TY_old.list$EP_LS.vec,1)*12/(month(LOSS_VAL_DATE)-month(TREATY_START)+1)*1.07,-5)
# new (CY)
EP_CY_new.list$EP_GC.vec*factor_GC_wp
EP_CY_new.list$EP_LS.vec*factor_LS_wp
# new (TY)
EP_TY_new.list$EP_GC.vec*factor_GC_wp+EP_TY_new.list$EP_LS.vec*factor_LS_wp
# new (TY original most recent year)
round(tail(EP_TY_old.list$EP_GC.vec+EP_TY_old.list$EP_LS.vec,1)*12/(month(LOSS_VAL_DATE)-month(TREATY_START)+1)*1.07,-5)
# new (TY revised most recent year)
round(tail(EP_TY_new.list$EP_GC.vec+EP_TY_new.list$EP_LS.vec,1)*12/(month(LOSS_VAL_DATE)-month(TREATY_START)+1),-5)
# new (TY project next year)
round(tail(EP_TY_new.list$EP_GC.vec+EP_TY_new.list$EP_LS.vec,1)*12/(month(LOSS_VAL_DATE)-month(TREATY_START)+1)*1.05,-5)



## In-force Portfolio summary

aggregate(df_policy_IF_new$Written_Premium,
          by=list(pmin(pmax(0,(df_policy_IF_new$Limit-1e06)/(5e05)),3),df_policy_IF_new$GC_LS),
          function(x)c(sum(x),length(x)))
aggregate(df_policy_IF_old$Written_Premium,
          by=list(pmin(pmax(0,(df_policy_IF_old$Limit-1e06)/(5e05)),3),df_policy_IF_old$GC_LS),
          function(x)c(sum(x),length(x)))


# ---------------------------
# Claims Simulation
# ---------------------------
claims <- list()

for (ay in AYs) {
  n_pol_ay <- sum(df_full$Policy_Year==ay)
  n_claim_ay <- n_claim[as.character(ay)]
  pol_sel.idx <- sample(n_pol_ay,size=n_claim_ay,replace=T)
  pol_sel <- df_full[df_full$Policy_Year==ay,][pol_sel.idx,]
  acc_dates <- pol_sel$Effective_Date + round(runif(n_claim_ay)*(pol_sel$Expiration_Date-pol_sel$Effective_Date))
  report_lags <- sample_report_lags(n_claim_ay)
  settle_lags <- sample_settle_lags(n_claim_ay)
  report_dates <- acc_dates + report_lags
  settle_dates <- report_dates + settle_lags
  
  time_ratio_paystart <- runif(n_claim_ay, min=0.01, max=0.25)
  time_ratio_paytrend1 <- runif(n_claim_ay, min=0.3, max=0.7)
  time_ratio_paytrend2 <- runif(n_claim_ay, min=0.4, max=0.6)
  time_ratio_payment <- pmax(pmin((as.numeric(LOSS_VAL_DATE-report_dates)-time_ratio_paystart*as.numeric(settle_dates-report_dates))/((1-time_ratio_paystart)*as.numeric(settle_dates-report_dates)),1),0)
  paid_ratio1 <- time_ratio_payment^time_ratio_paytrend1
  paid_ratio2 <- time_ratio_payment^time_ratio_paytrend2
  
  Ultimate_Loss <- pol_sel$Limit*pmin(0.05*prem_factor[as.character(ay)]*sample_severity(n_claim_ay)*(1+0.25*(pol_sel$State%in%c("CA","NY","IL","AZ")))*(1+0.25*(pol_sel$GC_LS=="GC")),1)
  ALAE_ratio <- runif(n_claim_ay, min=0.1, max=0.3)
  Ultimate_ALAE <- Ultimate_Loss*ALAE_ratio
  time_ratio_develop <- pmax(pmin((as.numeric(LOSS_VAL_DATE-report_dates))/as.numeric(settle_dates-report_dates),1),0)
  incurred_ratio_mean <- (1+time_ratio_develop^time_ratio_paytrend1)/2
  incurred_ratio_random1=rexp(n_claim_ay)
  incurred_ratio1 <- incurred_ratio_mean*(1*incurred_ratio_mean+incurred_ratio_random1*(1-incurred_ratio_mean))
  incurred_ratio_random2=rexp(n_claim_ay)
  incurred_ratio2 <- incurred_ratio_mean*(1*incurred_ratio_mean+incurred_ratio_random2*(1-incurred_ratio_mean))
  Incurred_Loss <- round(Ultimate_Loss*incurred_ratio1,2)
  Incurred_ALAE <- round(Ultimate_ALAE*incurred_ratio2,2)
  Paid_Loss <- round(Incurred_Loss*paid_ratio1,2)
  Paid_ALAE <- round(Incurred_ALAE*paid_ratio2,2)
  Outstanding_Loss <- Incurred_Loss - Paid_Loss
  Outstanding_ALAE <- Incurred_ALAE - Paid_ALAE
  Total_Incurred <- Incurred_Loss+Incurred_ALAE
  Status <- ifelse(settle_dates<=LOSS_VAL_DATE,"Closed","Open")
  
  claims[[as.character(ay)]]=pol_sel
  claims[[as.character(ay)]]$Claim_ID="temp"
  claims[[as.character(ay)]]$Status=Status
  claims[[as.character(ay)]]$Paid_Loss=Paid_Loss
  claims[[as.character(ay)]]$Outstanding_Loss=Outstanding_Loss
  claims[[as.character(ay)]]$Incurred_Loss=Incurred_Loss
  claims[[as.character(ay)]]$Paid_ALAE=Paid_ALAE
  claims[[as.character(ay)]]$Outstanding_ALAE=Outstanding_ALAE
  claims[[as.character(ay)]]$Incurred_ALAE=Incurred_ALAE
  claims[[as.character(ay)]]$Total_Incurred=Total_Incurred
  claims[[as.character(ay)]]$Ultimate_Loss=Ultimate_Loss
  claims[[as.character(ay)]]$Loss_Date=acc_dates
  claims[[as.character(ay)]]$Report_Date=report_dates
  claims[[as.character(ay)]]$Settle_Date=settle_dates
  claims[[as.character(ay)]]$time_ratio_paystart=time_ratio_paystart
  claims[[as.character(ay)]]$time_ratio_paytrend1=time_ratio_paytrend1
  claims[[as.character(ay)]]$time_ratio_paytrend2=time_ratio_paytrend2
  claims[[as.character(ay)]]$ALAE_ratio=ALAE_ratio
  claims[[as.character(ay)]]$incurred_ratio_random1=incurred_ratio_random1
  claims[[as.character(ay)]]$incurred_ratio_random2=incurred_ratio_random2
}
df_claim = bind_rows(claims)
df_claim = df_claim[df_claim$Report_Date<=LOSS_VAL_DATE,]
df_claim = df_claim[order(df_claim$Report_Date),]
df_claim$Claim_ID = sprintf(paste0("C-","-%05d"), 1:nrow(df_claim))
rownames(df_claim)=NULL

df_claim$AY=year(df_claim$Loss_Date)
df_claim_output=df_claim[,c("Claim_ID","Insured_ID","Loss_Date","AY","Policy_ID",
                            "Effective_Date","Policy_Year","Status","GC_LS","Limit",
                            "Paid_Loss","Outstanding_Loss","Incurred_Loss","Paid_ALAE",
                            "Outstanding_ALAE","Incurred_ALAE","Total_Incurred")]
df_claim_output_new=df_claim_output[(df_claim$Incurred_Loss>=10000)&(df_claim$AY<=year(LOSS_VAL_DATE))&(df_claim$AY>(year(LOSS_VAL_DATE)-5)),]
df_claim_output_old=df_claim_output[(df_claim$Incurred_Loss>=10000)&(df_claim$AY<=(year(LOSS_VAL_DATE)-1))&(df_claim$AY>(year(LOSS_VAL_DATE)-6))&(df_claim$Report_Date<=(LOSS_VAL_DATE-years(1))),]

# ---------------------------
# Development Triangles 
# ---------------------------
sim_triangles <- function(AYs_eval, df_claim_eval, loss_date_eval) {
  triangle_Incurred_Loss=triangle_Incurred_ALAE=array(NA,dim=c(length(AYs_eval),length(AYs_eval)))
  triangle_Paid_Loss=triangle_Paid_ALAE=array(NA,dim=c(length(AYs_eval),length(AYs_eval)))
  triangle_Report_Count=triangle_Settle_Count=array(NA,dim=c(length(AYs_eval),length(AYs_eval)))
  for (ay in AYs_eval) {
    idx_ay=match(ay,AYs_eval)
    for (dev in 1:(length(AYs_eval)-idx_ay+1)) {
      eval_date=make_date(ay+dev-1, month(loss_date_eval), day(loss_date_eval))
      claim_sel=df_claim_eval[(year(df_claim_eval$Loss_Date)==ay)&(df_claim_eval$Report_Date<=eval_date),]
      triangle_Report_Count[idx_ay,dev]=sum(claim_sel$Report_Date<=eval_date)
      triangle_Settle_Count[idx_ay,dev]=sum(claim_sel$Settle_Date<=eval_date)
      
      report_dates=claim_sel$Report_Date
      settle_dates=claim_sel$Settle_Date
      time_ratio_paystart=claim_sel$time_ratio_paystart
      time_ratio_paytrend1=claim_sel$time_ratio_paytrend1
      time_ratio_paytrend2=claim_sel$time_ratio_paytrend2
      ALAE_ratio=claim_sel$ALAE_ratio
      incurred_ratio_random1=claim_sel$incurred_ratio_random1
      incurred_ratio_random2=claim_sel$incurred_ratio_random2
      Ultimate_Loss=claim_sel$Ultimate_Loss
      
      time_ratio_payment <- pmax(pmin((as.numeric(eval_date-report_dates)-time_ratio_paystart*as.numeric(settle_dates-report_dates))/((1-time_ratio_paystart)*as.numeric(settle_dates-report_dates)),1),0)
      paid_ratio1 <- time_ratio_payment^time_ratio_paytrend1
      paid_ratio2 <- time_ratio_payment^time_ratio_paytrend2
      Ultimate_ALAE <- Ultimate_Loss*ALAE_ratio
      time_ratio_develop <- pmax(pmin((as.numeric(eval_date-report_dates))/as.numeric(settle_dates-report_dates),1),0)
      incurred_ratio_mean <- (1+time_ratio_develop^time_ratio_paytrend1)/2
      incurred_ratio1 <- incurred_ratio_mean*(1*incurred_ratio_mean+incurred_ratio_random1*(1-incurred_ratio_mean))
      incurred_ratio2 <- incurred_ratio_mean*(1*incurred_ratio_mean+incurred_ratio_random2*(1-incurred_ratio_mean))
      Incurred_Loss <- round(Ultimate_Loss*incurred_ratio1,2)
      Incurred_ALAE <- round(Ultimate_ALAE*incurred_ratio2,2)
      Paid_Loss <- round(Incurred_Loss*paid_ratio1,2)
      Paid_ALAE <- round(Incurred_ALAE*paid_ratio2,2)
      triangle_Incurred_Loss[idx_ay,dev]=sum(Incurred_Loss)
      triangle_Incurred_ALAE[idx_ay,dev]=sum(Incurred_ALAE)
      triangle_Paid_Loss[idx_ay,dev]=sum(Paid_Loss)
      triangle_Paid_ALAE[idx_ay,dev]=sum(Paid_ALAE)
    }
  }
  triangle_Incurred_Total=triangle_Incurred_Loss+triangle_Incurred_ALAE
  triangle_Paid_Total=triangle_Paid_Loss+triangle_Paid_ALAE
  list(triangle_Report_Count=triangle_Report_Count,
       triangle_Settle_Count=triangle_Settle_Count,
       triangle_Incurred_Loss=triangle_Incurred_Loss,
       triangle_Incurred_ALAE=triangle_Incurred_ALAE,
       triangle_Incurred_Total=triangle_Incurred_Total, 
       triangle_Paid_Loss=triangle_Paid_Loss,
       triangle_Paid_ALAE=triangle_Paid_ALAE,
       triangle_Paid_Total=triangle_Paid_Total)
}

sim_triangles(AYs_eval=AYs, df_claim_eval=df_claim, loss_date_eval=LOSS_VAL_DATE)

########### Build with earlier AY ###########
AYs_extract <- AYs[-length(AYs)]
AYs_extract_char <- as.character(AYs_extract)
LOSS_VAL_DATE_extract = eval_date = make_date(tail(AYs_extract,1), month(LOSS_VAL_DATE), day(LOSS_VAL_DATE))
df_full_extract <- df_full[df_full$Effective_Date<=LOSS_VAL_DATE_extract,]
df_claim_extract <- df_claim[df_claim$Report_Date<=LOSS_VAL_DATE_extract,]

claim_sel = df_claim_extract
report_dates=claim_sel$Report_Date
settle_dates=claim_sel$Settle_Date
time_ratio_paystart=claim_sel$time_ratio_paystart
time_ratio_paytrend1=claim_sel$time_ratio_paytrend1
time_ratio_paytrend2=claim_sel$time_ratio_paytrend2
ALAE_ratio=claim_sel$ALAE_ratio
incurred_ratio_random1=claim_sel$incurred_ratio_random1
incurred_ratio_random2=claim_sel$incurred_ratio_random2
Ultimate_Loss=claim_sel$Ultimate_Loss

time_ratio_payment <- pmax(pmin((as.numeric(eval_date-report_dates)-time_ratio_paystart*as.numeric(settle_dates-report_dates))/((1-time_ratio_paystart)*as.numeric(settle_dates-report_dates)),1),0)
paid_ratio1 <- time_ratio_payment^time_ratio_paytrend1
paid_ratio2 <- time_ratio_payment^time_ratio_paytrend2
Ultimate_ALAE <- Ultimate_Loss*ALAE_ratio
time_ratio_develop <- pmax(pmin((as.numeric(eval_date-report_dates))/as.numeric(settle_dates-report_dates),1),0)
incurred_ratio_mean <- (1+time_ratio_develop^time_ratio_paytrend1)/2
incurred_ratio1 <- incurred_ratio_mean*(1*incurred_ratio_mean+incurred_ratio_random1*(1-incurred_ratio_mean))
incurred_ratio2 <- incurred_ratio_mean*(1*incurred_ratio_mean+incurred_ratio_random2*(1-incurred_ratio_mean))
Incurred_Loss <- round(Ultimate_Loss*incurred_ratio1,2)
Incurred_ALAE <- round(Ultimate_ALAE*incurred_ratio2,2)
Paid_Loss <- round(Incurred_Loss*paid_ratio1,2)
Paid_ALAE <- round(Incurred_ALAE*paid_ratio2,2)
Outstanding_Loss <- Incurred_Loss - Paid_Loss
Outstanding_ALAE <- Incurred_ALAE - Paid_ALAE
Total_Incurred <- Incurred_Loss+Incurred_ALAE
Status <- ifelse(settle_dates<=eval_date,"Closed","Open")

df_claim_extract$Status=Status
df_claim_extract$Paid_Loss=Paid_Loss
df_claim_extract$Outstanding_Loss=Outstanding_Loss
df_claim_extract$Incurred_Loss=Incurred_Loss
df_claim_extract$Paid_ALAE=Paid_ALAE
df_claim_extract$Outstanding_ALAE=Outstanding_ALAE
df_claim_extract$Incurred_ALAE=Incurred_ALAE
df_claim_extract$Total_Incurred=Total_Incurred

sim_triangles(AYs_eval=AYs_extract, df_claim_eval=df_claim_extract, loss_date_eval=LOSS_VAL_DATE_extract)



############ Output in CSV ##############
write.csv(df_full, file = "Sample1_QS_policy.csv", row.names = FALSE)
write.csv(df_policy_IF_new, file = "Sample1_QS_policy_IF_new.csv", row.names = FALSE)
write.csv(df_policy_IF_old, file = "Sample1_QS_policy_IF_old.csv", row.names = FALSE)
write.csv(df_claim_output, file = "Sample1_QS_claim.csv", row.names = FALSE)
write.csv(df_claim_output_new, file = "Sample1_QS_claim_new.csv", row.names = FALSE)
write.csv(df_claim_output_old, file = "Sample1_QS_claim_old.csv", row.names = FALSE)