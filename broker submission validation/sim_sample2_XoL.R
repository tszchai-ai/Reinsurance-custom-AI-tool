# ============================================
# Fake Broker Submission Generator
# ============================================
library(lubridate)
library(openxlsx)
library(dplyr)
library(VGAM)

set.seed(20260131)  # reproducible, but all figures are random draws

# ---------------------------
# Global Parameters / Dates
# ---------------------------
LOSS_VAL_DATE  <- as.Date("2024-12-31")
EXPO_ASOF_DATE <- LOSS_VAL_DATE
TREATY_START   <- as.Date("2025-07-01")
TREATY_END     <- as.Date("2026-06-30")
AYs            <- 2015:2024
AYs_char       <- as.character(AYs)


##### Parameters for model input
## Poisson means for policy count and claim count
# lambda for policy
lambda_pol_1 <- c(800, 800, 800, 900, 900, 900, 1000, 1000, 1100, 1200)
lambda_pol_2 <- c(0,  0,    0,   0,   0,   300,  500,  800, 1000, 1300)
lambda_pol_3 <- c(1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000)
n_pol_1 <- rpois(length(AYs_char),lambda=lambda_pol_1)
n_pol_2 <- rpois(length(AYs_char),lambda=lambda_pol_2)
n_pol_3 <- rpois(length(AYs_char),lambda=lambda_pol_3)
names(n_pol_1)=names(n_pol_2)=names(n_pol_3)=AYs_char
goforward_n <- rpois(3,lambda=c(1300,1500,1000))
# lambda for avg claim per policy
lambda_claim_1 <- c(0.05,0.05,0.055,0.055,0.065,0.065,0.07,0.07,0.06,0.055)
lambda_claim_2 <- c(0,0,0,0,0,0.03,0.04,0.05,0.06,0.07)
lambda_claim_3 <- c(0.03,0.03,0.03,0.03,0.03,0.03,0.03,0.03,0.03,0.03)
n_claim_1 <- rpois(length(AYs_char),lambda=(n_pol_1*lambda_claim_1))
n_claim_2 <- rpois(length(AYs_char),lambda=(n_pol_2*lambda_claim_2))
n_claim_3 <- rpois(length(AYs_char),lambda=(n_pol_3*lambda_claim_3))
names(n_claim_1)=names(n_claim_2)=names(n_claim_3)=AYs_char
names(lambda_claim_1)=names(lambda_claim_2)=names(lambda_claim_3)=AYs_char
# renewal probability
prob_renew <- c(0.00, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85)
n_renew_1 <- c(0,rbinom(length(AYs_char)-1,size=n_pol_1[-length(AYs_char)],prob=prob_renew[-1]))
n_renew_2 <- c(0,rbinom(length(AYs_char)-1,size=n_pol_2[-length(AYs_char)],prob=prob_renew[-1]))
n_renew_3 <- c(0,rbinom(length(AYs_char)-1,size=n_pol_3[-length(AYs_char)],prob=prob_renew[-1]))
names(n_renew_1)=names(n_renew_2)=names(n_renew_3)=names(prob_renew)=AYs_char
# states
states   <- c("AK","AL","AR","AZ","CA","CO","CT","DC","DE","FL",
              "GA","IA","ID","IL","IN","KS","KY","LA","MA","MD",
              "ME","MI","MN","MO","MS","MT","NC","ND","NE","NH",
              "NJ","NM","NV","NY","OH","OK","OR","PA","RI","SC",
              "SD","TN","TX","UT","VA","VT","WA","WI","WV","WY")
states_gdp <- c(AK=70.0,  AL=321.2, AR=188.7, AZ=552.2, CA=4103.1,
                CO=553.3, CT=365.7, DC=74.0, DE=103.3, FL=1705.6, 
                GA=882.5, IA=257.0, ID=128.1, IL=1137.2, IN=527.4,
                KS=234.7, KY=293.0, LA=327.8, MA=780.7, MD=542.8,
                ME=98.6,  MI=706.6, MN=500.9, MO=451.2, MS=157.5,
                MT=76.0,  NC=839.1, ND=75.4,  NE=185.4, NH=121.2,
                NJ=846.6, NM=140.5, NV=260.7, NY=2297.0, OH=927.7,
                OK=265.8, OR=331.0, PA=1024.2, RI=82.5,  SC=350.0,
                SD=75.2,  TN=549.7, TX=2709.4, UT=300.9, VA=764.5,
                VT=45.7,  WA=854.7, WI=451.3, WV=107.7, WY=52.9)
states_pop <- c(AK=0.737,  AL=5.193, AR=3.115,  AZ=7.624,  CA=39.355,
                CO=6.013,  CT=3.688,  DC=0.700 , DE=1.060,  FL=23.463,  
                GA=11.303,  IA=3.238,  ID=2.030,  IL=12.719, IN=6.973,  
                KS=2.977,  KY=4.607,  LA=4.618,  MA=7.154, MD=6.265,  
                ME=1.415,  MI=10.128,  MN=5.830,  MO=6.271, MS=2.954,  
                MT=1.145,  NC=11.198,  ND=0.799,  NE=2.018, NH=1.415,  
                NJ=9.548,  NM=2.125,  NV=3.282,  NY=20.002, OH=11.901,  
                OK=4.123,  OR=4.274,  PA=13.059,  RI=1.115, SC=5.570, 
                SD=0.935,  TN=7.315,  TX=31.710,  UT=3.539, VA=8.880,  
                VT=0.645,  WA=8.001,  WI=5.973,  WV=1.766, WY=0.589)
states_prob <- states_pop/sum(states_pop)
state_prem_ratio <- (states_gdp/states_pop)/mean(states_gdp/states_pop)
# Limits 
limit_options <- c(1e6, 2e6, 3e6, 4e6, 5e6, 1e7)
limit_probs_1 <- c(0.37, 0.18, 0.10, 0.05, 0.25, 0.05)
limit_probs_2 <- c(0.66, 0.12, 0.05, 0.05, 0.07, 0.05)
limit_probs_3 <- c(0.42, 0.19, 0.05, 0.05, 0.24, 0.05)

# Subline 1: Commercial Auto Umbrella – Fleet
# 50 descriptions, ordered least severe -> most severe (10 words max each)
desc_1 <- c(
  "Minor rear-end; soft tissue; quick settlement",
  "Low-speed collision; neck strain; brief PT",
  "Parking lot backing crash; minor BI alleged",
  "Sideswipe on arterial; chiropractic treatment claimed",
  "Minor lane-change contact; ER visit; released",
  "Delivery van fender-bender; passenger complaint filed",
  "Minor intersection bump; whiplash; meds only",
  "Bumper-to-bumper crash; short wage loss claimed",
  "Rear-end at light; disc bulge allegation",
  "Fleet pickup hit curb; pedestrian stumble claim",
  "Two-vehicle crash; moderate BI; lawsuit filed",
  "Left-turn collision; fractured wrist; surgery consult",
  "Merge collision; multiple claimants; limits tendered",
  "Truck struck stopped car; concussion diagnosis",
  "Backing tractor-trailer; loading dock injury claim",
  "Multi-car chain reaction; several BI claimants",
  "Delivery truck ran red; significant BI litigation",
  "Rollover single-vehicle; passenger serious injuries",
  "Commercial van T-bone; hospitalized claimant",
  "Fleet SUV hit cyclist; multiple fractures",
  "Tractor-trailer jackknife; injuries to motorists",
  "Box truck struck pedestrian; severe leg injury",
  "Fleet vehicle crossed centerline; head-on impact",
  "Highway pileup involving fleet; multiple ambulances",
  "Truck rear-ended bus; many injury allegations",
  "Fuel spill after crash; evacuation; BI claims",
  "Van struck motorcycle; critical injuries reported",
  "Fleet truck hit work zone; worker injured",
  "Commercial auto hit school bus; injuries alleged",
  "Tractor-trailer underride; catastrophic injuries claimed",
  "Fatal crash; wrongful death claim filed",
  "Multi-fatality highway crash; extensive litigation",
  "Pedestrian fatality; surveillance disputes; trial set",
  "Truck struck family vehicle; two fatalities",
  "Bus crash; multiple severe injuries; media attention",
  "Commercial auto hit bridge; passenger fatalities",
  "Hazmat truck crash; exposure claims; class counsel",
  "Fleet vehicle hit stroller; infant critical injuries",
  "Severe TBI claim; long-term care demanded",
  "Quadriplegia allegation after crash; life care plan",
  "Burn injuries from post-crash fire; complex damages",
  "Amputation claim; disputed liability; expert retained",
  "Spinal cord injury; lifetime medical costs alleged",
  "Wrongful death with punitive demand; nuclear verdict risk",
  "Multiple wrongful deaths; consolidated litigation filed",
  "Catastrophic bus rollover; dozens injured; high exposure",
  "Fleet driver DUI allegation; punitive damages sought",
  "Nuclear verdict risk; severe BI; venue unfavorable",
  "Excess exposure; multiple catastrophics; global settlement"
)

# Subline 2: Commercial Auto Umbrella – Non-Fleet
desc_2 <- c(
  "Minor rear-end; claimant sore; no treatment documented",
  "Low-speed backing accident; small BI demand",
  "Parking lot collision; minor injury alleged",
  "Side-swipe while merging; chiropractic visits claimed",
  "Intersection bump; ER visit; discharged same day",
  "Minor rear-end; missed work days claimed",
  "Lane change contact; soft tissue; counsel retained",
  "Left-turn collision; moderate BI; settlement talks",
  "Rear-end with passenger; disc injury alleged",
  "Backing into pedestrian; bruising; short PT",
  "Two-car crash; fractures; surgery recommended",
  "Delivery-type vehicle; passenger injured; suit filed",
  "Distracted driving allegation; BI demand escalated",
  "T-bone at intersection; hospitalization; lien asserted",
  "Non-fleet pickup hit cyclist; multiple fractures",
  "Single-vehicle rollover; serious passenger injuries",
  "Commercial van hit tree; severe BI claim",
  "Head-on collision; significant BI; liability disputed",
  "Multi-vehicle crash; several BI claimants",
  "Motorcycle collision; critical injuries reported",
  "Pedestrian struck in crosswalk; severe leg injury",
  "Work zone crash; worker injured; OSHA inquiry",
  "Company vehicle hit bus; multiple injury claims",
  "Underride scenario; catastrophic injury allegations",
  "Crash with elderly driver; fatality; wrongful death",
  "Fatal pedestrian claim; high economic damages",
  "Severe TBI allegation; future care demanded",
  "Spinal injury claimed; long-term disability alleged",
  "Amputation allegation; expert reports exchanged",
  "Burn injuries from collision; extensive treatment",
  "Wrongful death; punitive allegations; discovery disputes",
  "Two fatalities; consolidated wrongful death suits",
  "Multi-claimant litigation; joint settlement negotiation",
  "Catastrophic BI; life care plan presented",
  "Permanent paralysis claim; lifetime expenses sought",
  "High-venue plaintiff counsel; aggressive trial posture",
  "Reptile theory; corporate negligence alleged",
  "Cell-phone distraction; punitive damages pursued",
  "Severe disfigurement claim; damages escalating",
  "Excess limits demand; time-limit letter received",
  "Umbrella exposure; mediation failed; trial scheduled",
  "Verdict risk; sympathetic claimant; adverse venue",
  "Multiple catastrophics; layered settlement structure",
  "Wrongful death plus punitive; bad faith threatened",
  "Nuclear verdict warning; large reserve increase",
  "Catastrophic spinal cord; structured settlement discussed",
  "Severe TBI; guardian appointed; lifetime care",
  "Excess exposure; multiple claimants; global resolution",
  "Policy limits tendered; umbrella fully exposed"
)

# Subline 3: Commercial Auto Umbrella – Specialty (Public Auto / Livery / Specialty)
desc_3 <- c(
  "Minor passenger slip on bus step; medical visit",
  "Low-speed shuttle bump; passenger soreness claimed",
  "Taxi rear-end; soft tissue; small demand",
  "Paratransit abrupt stop; passenger fall alleged",
  "Livery sideswipe; minor BI; treatment minimal",
  "Bus mirror strike; pedestrian complaint filed",
  "Passenger dispute; minor assault allegation; injury claimed",
  "Shuttle lane change; passenger neck strain",
  "Taxi collision; claimant attorney retained quickly",
  "Bus hard brake; multiple passenger complaints",
  "Paratransit wheelchair tip; moderate injury alleged",
  "School transport incident; parent claim; counsel involved",
  "Shuttle crash; several passengers treated",
  "Bus rear-end; multiple BI claimants; litigation",
  "Taxi hit cyclist; fractures; liability contested",
  "Livery vehicle hit pedestrian; severe leg injury",
  "Paratransit crash; vulnerable passenger injured",
  "Bus door malfunction; passenger trapped; injury claim",
  "Shuttle rollover; passengers hospitalized",
  "Charter bus collision; multiple serious injuries",
  "Bus struck vehicle; passenger ejected; critical injuries",
  "Public auto crash; injuries plus ADA compliance allegations",
  "Bus fire after crash; smoke inhalation claims",
  "Taxi hit motorcycle; catastrophic injury reported",
  "Shuttle hit work zone; worker injured",
  "Bus crash with fatality; wrongful death filed",
  "Multiple passenger fatalities; mass tort counsel engaged",
  "Catastrophic bus rollover; dozens injured; high exposure",
  "Severe TBI passenger claim; extensive future care",
  "Spinal cord injury passenger; lifetime support demanded",
  "Wheelchair passenger paralysis allegation; high damages",
  "Burn injuries in bus fire; complex medical course",
  "Amputation after livery crash; experts retained",
  "Wrongful death with punitive; driver conduct challenged",
  "Driver fatigue allegation; punitive damages pursued",
  "DUI allegation; public carrier; reputational risk",
  "Negligent hiring claim; prior violations alleged",
  "Video evidence dispute; passengers allege reckless driving",
  "Multi-venue suits; consolidated discovery ordered",
  "Class counsel claims systemic safety failures",
  "Nuclear verdict risk; sympathetic passengers; adverse venue",
  "Mass casualty incident; global settlement discussions",
  "Multiple wrongful deaths; layered umbrella exhausted",
  "Severe disfigurement; emotional distress claims added",
  "High economic loss; key earner fatality claimed",
  "Catastrophic brain injury; guardianship and liens",
  "Punitive exposure; safety policy violations alleged",
  "Excess limits demand; mediation impasse; trial looming",
  "Umbrella fully exposed; multiple catastrophics; final settlement"
)

## Rate change
# Subline 1
rate_change_avg_new <- round(runif(length(AYs_char), -3, 10), 1)
rate_change_avg_renew <- rate_change_avg_new*runif(length(AYs_char), 0.5, 0.8)
names(rate_change_avg_new) = names(rate_change_avg_renew) = AYs_char
df <- expand.grid(State = states, Year = AYs, KEEP.OUT.ATTRS = FALSE, stringsAsFactors = FALSE)
df$addl_change <- runif(nrow(df), -2, 2)
df$`Rate change (New)` <- rate_change_avg_new[as.character(df$Year)] + df$addl_change
df$`Rate change (Renew)` <- rate_change_avg_renew[as.character(df$Year)] + df$addl_change
rate_change_1 <- df[, c("State", "Year", "Rate change (New)", "Rate change (Renew)")]
rate_change_1$`Rate change (New)` <- round(rate_change_1$`Rate change (New)`, 1)
rate_change_1$`Rate change (Renew)` <- round(rate_change_1$`Rate change (Renew)`, 1)
ratio_renew = runif(nrow(df), 0.6, 0.8)
rate_change_1$`Rate change`<- round(rate_change_1$`Rate change (New)`*(1-ratio_renew)+rate_change_1$`Rate change (Renew)`*ratio_renew, 1)
rate_mat_1 = with(rate_change_1, tapply(`Rate change`, list(State, Year), FUN = identity))
factor_mat_1 = t(round(apply(rate_mat_1,1,function(x)c(1,cumprod(1+x/100))),3))
prem_rate_1=round(rate_change_avg_new*(1-mean(ratio_renew))+rate_change_avg_renew*mean(ratio_renew),1)
prem_factor_1=round(c(1,cumprod(1+prem_rate_1/100)),3)
# Subline 2
rate_change_avg_new <- round(runif(length(AYs_char), -3, 18), 1)
rate_change_avg_renew <- rate_change_avg_new*runif(length(AYs_char), 0.7, 1)
names(rate_change_avg_new) = names(rate_change_avg_renew) = AYs_char
df <- expand.grid(State = states, Year = AYs, KEEP.OUT.ATTRS = FALSE, stringsAsFactors = FALSE)
df$addl_change <- runif(nrow(df), -3, 3)
df$`Rate change (New)` <- rate_change_avg_new[as.character(df$Year)] + df$addl_change
df$`Rate change (Renew)` <- rate_change_avg_renew[as.character(df$Year)] + df$addl_change
rate_change_2 <- df[, c("State", "Year", "Rate change (New)", "Rate change (Renew)")]
rate_change_2$`Rate change (New)` <- round(rate_change_2$`Rate change (New)`, 1)
rate_change_2$`Rate change (Renew)` <- round(rate_change_2$`Rate change (Renew)`, 1)
ratio_renew = runif(nrow(df), 0.6, 0.8)
rate_change_2$`Rate change`<- round(rate_change_2$`Rate change (New)`*(1-ratio_renew)+rate_change_2$`Rate change (Renew)`*ratio_renew, 1)
rate_mat_2 = with(rate_change_2, tapply(`Rate change`, list(State, Year), FUN = identity))
factor_mat_2 = t(round(apply(rate_mat_2,1,function(x)c(1,cumprod(1+x/100))),3))
prem_rate_2=round(rate_change_avg_new*(1-mean(ratio_renew))+rate_change_avg_renew*mean(ratio_renew),1)
prem_factor_2=round(c(1,cumprod(1+prem_rate_2/100)),3)
# Subline 3
rate_change_avg_new <- round(runif(length(AYs_char), -2, 12), 1)
rate_change_avg_renew <- rate_change_avg_new*runif(length(AYs_char), 0.6, 1)
names(rate_change_avg_new) = names(rate_change_avg_renew) = AYs_char
df <- expand.grid(State = states, Year = AYs, KEEP.OUT.ATTRS = FALSE, stringsAsFactors = FALSE)
df$addl_change <- runif(nrow(df), -2, 2)
df$`Rate change (New)` <- rate_change_avg_new[as.character(df$Year)] + df$addl_change
df$`Rate change (Renew)` <- rate_change_avg_renew[as.character(df$Year)] + df$addl_change
rate_change_3 <- df[, c("State", "Year", "Rate change (New)", "Rate change (Renew)")]
rate_change_3$`Rate change (New)` <- round(rate_change_3$`Rate change (New)`, 1)
rate_change_3$`Rate change (Renew)` <- round(rate_change_3$`Rate change (Renew)`, 1)
ratio_renew = runif(nrow(df), 0.6, 0.8)
rate_change_3$`Rate change`<- round(rate_change_3$`Rate change (New)`*(1-ratio_renew)+rate_change_3$`Rate change (Renew)`*ratio_renew, 1)
rate_mat_3 = with(rate_change_3, tapply(`Rate change`, list(State, Year), FUN = identity))
factor_mat_3 = t(round(apply(rate_mat_3,1,function(x)c(1,cumprod(1+x/100))),3))
prem_rate_3=round(rate_change_avg_new*(1-mean(ratio_renew))+rate_change_avg_renew*mean(ratio_renew),1)
prem_factor_3=round(c(1,cumprod(1+prem_rate_3/100)),3)

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
factor_wp_1=c(1,1,1,1,1,1,1.0011,1.0049,1.016,1)
factor_wp_2=c(1,1,1,1,1,1,1,1,1.008,1)
factor_wp_3=c(1,1,1,1,1,1,1.0008,1.0037,1.012,1)

# ---------------------------
# Utility Functions
# ---------------------------
rand_date_between <- function(start, end, n) {
  start + sample.int(as.integer(end - start) + 1L, n, replace = TRUE) - 1L
}

# ---------------------------
# Build Exposure information: Output df_full.
# ---------------------------
sim_exposure <- function(n_pol, n_renew, limit_probs, factor_mat, factor_wp,
                         prem_factor, prem_scale, line.idx)
{
  df.list=df_new.list=df_renew.list=list()
  for (ay in AYs[which(n_pol>0)]) {
    n_pol_new_ay=max(n_pol[as.character(ay)]-n_renew[as.character(ay)],10)
    df.eff.date=rand_date_between(make_date(ay, 1, 1), make_date(ay, 12, 31), n_pol_new_ay)
    df.state=sample(states, n_pol_new_ay, replace = TRUE, prob=states_prob)
    df.limit=sample(limit_options, n_pol_new_ay, replace = TRUE, prob = limit_probs)
    df.status="New"
    df.loading=factor_mat[df.state,as.character(ay)]*state_prem_ratio[df.state]*runif(n_pol_new_ay,min=0.9,max=1.1)
    df.written.prem=round(prem_scale*df.loading*df.limit,2)
    df_new.list[[as.character(ay)]] <- data.frame(
      Policy_ID = "temp", #sprintf(paste0("P-",ay,"-%05d"), 1:n_pol_new_ay),
      Insured_ID = "temp", #sprintf(paste0("I-",ay,"-%05d"), 1:n_pol_new_ay),
      Policy_Year = ay,
      Effective_Date  = df.eff.date,
      Expiration_Date = df.eff.date%m+%years(1)-1,
      New_Renewed = df.status,
      State     = df.state,
      Written_Premium = df.written.prem,
      Limit     = df.limit,
      stringsAsFactors = FALSE)
    df_new.list[[as.character(ay)]]=df_new.list[[as.character(ay)]][order(df_new.list[[as.character(ay)]]$Effective_Date),]
    df_new.list[[as.character(ay)]]$Insured_ID=sprintf(paste0(ay-2000,"-",line.idx,"-%05d"), 1:n_pol_new_ay)
    if(ay>min(AYs[which(n_pol>0)])){
      n_pol_renew_ay=min(n_renew[as.character(ay)],n_pol[as.character(ay-1)]-10)
      df_renew.list[[as.character(ay)]]=df.list[[as.character(ay-1)]][sample(nrow(df.list[[as.character(ay-1)]]),n_pol_renew_ay,replace=F),]
      df_renew.list[[as.character(ay)]]$Effective_Date=df_renew.list[[as.character(ay)]]$Effective_Date%m+%years(1)
      df_renew.list[[as.character(ay)]]$Expiration_Date=df_renew.list[[as.character(ay)]]$Expiration_Date%m+%years(1)
      df_renew.list[[as.character(ay)]]$Policy_Year=df_renew.list[[as.character(ay)]]$Policy_Year+1
      df_renew.list[[as.character(ay)]]$New_Renewed="Renewed"
      df.loading=factor_mat[df_renew.list[[as.character(ay)]]$State,as.character(ay)]*state_prem_ratio[df_renew.list[[as.character(ay)]]$State]*runif(n_pol_renew_ay,min=0.9,max=1.1)
      df_renew.list[[as.character(ay)]]$Written_Premium=round(prem_scale*df.loading*df_renew.list[[as.character(ay)]]$Limit,2)
      df.list[[as.character(ay)]]=rbind(df_new.list[[as.character(ay)]],df_renew.list[[as.character(ay)]])
      df.list[[as.character(ay)]]=df.list[[as.character(ay)]][order(df.list[[as.character(ay)]]$Effective_Date),]
      df.list[[as.character(ay)]]$Policy_ID=sprintf(paste0("P",ay-2000,"-",line.idx,"-%05d"), 1:(n_pol_new_ay+n_pol_renew_ay))
      df_full=rbind(df_full,df.list[[as.character(ay)]])
    } else {
      df.list[[as.character(ay)]]=df_new.list[[as.character(ay)]]
      df.list[[as.character(ay)]]$Policy_ID=sprintf(paste0("P",ay-2000,"-",line.idx,"-%05d"), 1:n_pol_new_ay)
      df_full=df.list[[as.character(ay)]]
    }
  }
  df_full=df_full[df_full$Effective_Date<=LOSS_VAL_DATE,]
  rownames(df_full)=NULL
  
  df_policy_IF_new=df_full[(df_full$Effective_Date<=LOSS_VAL_DATE)&(df_full$Effective_Date>(LOSS_VAL_DATE-years(1))),]
  df_policy_IF_old=df_full[(df_full$Effective_Date<=(LOSS_VAL_DATE-years(1)))&(df_full$Effective_Date>(LOSS_VAL_DATE-years(2))),]
  
  limit_profile_new=aggregate(df_policy_IF_new$Written_Premium,by=list(df_policy_IF_new$Limit),FUN=function(x)c(sum(x),length(x)))
  limit_profile_old=aggregate(df_policy_IF_old$Written_Premium,by=list(df_policy_IF_old$Limit),FUN=function(x)c(sum(x),length(x)))
  df_out=aggregate(df_full$Written_Premium,by=list(df_full$State,df_full$Policy_Year),FUN=sum)
  wp_state_new=sweep(with(df_out, tapply(x, list(factor(Group.1, levels=states), Group.2), FUN = identity)),2,factor_wp[which(n_pol>0)],"*")
  wp_state_new[is.na(wp_state_new)]=0
  df_out=aggregate(df_full$Written_Premium[df_full$Effective_Date<=(LOSS_VAL_DATE-years(1))],
                   by=list(df_full$State[df_full$Effective_Date<=(LOSS_VAL_DATE-years(1))],
                           df_full$Policy_Year[df_full$Effective_Date<=(LOSS_VAL_DATE-years(1))]),FUN=sum)
  wp_state_old=with(df_out, tapply(x, list(factor(Group.1, levels=states), Group.2), FUN = identity))
  wp_state_old[is.na(wp_state_old)]=0
  
  list(df_full=df_full, df_policy_IF_new=df_policy_IF_new, df_policy_IF_old=df_policy_IF_old,
       limit_profile_new=limit_profile_new, limit_profile_old=limit_profile_old,
       wp_state_new=wp_state_new, wp_state_old=wp_state_old)
}

sim_expo_list_1=sim_exposure(n_pol_1, n_renew_1, limit_probs_1, factor_mat_1, factor_wp_1,
                             prem_factor_1, prem_scale=0.004, line.idx=1)
sim_expo_list_2=sim_exposure(n_pol_2, n_renew_2, limit_probs_2, factor_mat_2, factor_wp_2,
                             prem_factor_2, prem_scale=0.004, line.idx=2)
sim_expo_list_3=sim_exposure(n_pol_3, n_renew_3, limit_probs_3, factor_mat_3, factor_wp_3,
                             prem_factor_3, prem_scale=0.004, line.idx=3)
# ---------------------------
# Claims Simulation
# ---------------------------
sim_claim = function(df_full, n_claim, prem_factor, desc, claim_scale=0.05, line.idx=1, threshold_report=50000){
  claims <- list()
  for (ay in AYs[which(n_claim>0)]) {
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
    
    Ultimate_Loss <- pol_sel$Limit*claim_scale*prem_factor[as.character(ay)]*sample_severity(n_claim_ay)*state_prem_ratio[pol_sel$State]
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
  df_claim$Claim_ID = sprintf(paste0("C",line.idx,"-","-%05d"), 1:nrow(df_claim))
  rownames(df_claim)=NULL
  
  df_claim$AY=year(df_claim$Loss_Date)
  df_claim$Reserve=df_claim$Outstanding_Loss+df_claim$Outstanding_ALAE
  portion_ANO=runif(nrow(df_claim),min=0.2,max=0.3)
  df_claim$portion_ANO=portion_ANO
  df_claim$ANO=df_claim$Paid_ALAE*df_claim$portion_ANO
  df_claim$DCC=df_claim$Paid_ALAE*(1-df_claim$portion_ANO)
  df_claim$Description=desc[pmax(1,round(length(desc)*rank(df_claim$Ultimate_Loss)/length(df_claim$Ultimate_Loss)))]
  
  ## old year
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
  df_claim_extract$AY=year(df_claim_extract$Loss_Date)
  df_claim_extract$Reserve=df_claim_extract$Outstanding_Loss+df_claim_extract$Outstanding_ALAE
  df_claim_extract$ANO=df_claim_extract$Paid_ALAE*df_claim_extract$portion_ANO
  df_claim_extract$DCC=df_claim_extract$Paid_ALAE*(1-df_claim_extract$portion_ANO)
  
  df_claim_output=df_claim[,c("Claim_ID","Policy_ID","Insured_ID","Status","Effective_Date","State","Limit","Loss_Date","AY",
                              "ANO","DCC","Paid_Loss","Reserve","Total_Incurred","Description")]
  df_claim_output_new=df_claim_output[(df_claim$Incurred_Loss>=threshold_report)&(df_claim$AY<=year(LOSS_VAL_DATE))&(df_claim$AY>(year(LOSS_VAL_DATE)-5)),]
  
  df_claim_output=df_claim_extract[,c("Claim_ID","Policy_ID","Insured_ID","Status","Effective_Date","State","Limit","Loss_Date","AY",
                              "ANO","DCC","Paid_Loss","Reserve","Total_Incurred","Description")]
  df_claim_output_old=df_claim_output[(df_claim$Incurred_Loss>=threshold_report)&(df_claim$AY<=(year(LOSS_VAL_DATE)-1))&(df_claim$AY>(year(LOSS_VAL_DATE)-6))&(df_claim$Report_Date<=(LOSS_VAL_DATE-years(1))),]
  
  list(df_claim=df_claim,df_claim_output_new=df_claim_output_new,df_claim_output_old=df_claim_output_old)
}

sim_claim_list_1=sim_claim(sim_expo_list_1$df_full, n_claim_1, prem_factor_1, desc=desc_1, claim_scale=0.05, line.idx=1, threshold_report=50000)
sim_claim_list_2=sim_claim(sim_expo_list_2$df_full, n_claim_2, prem_factor_3, desc=desc_2, claim_scale=0.05, line.idx=1, threshold_report=50000)
sim_claim_list_3=sim_claim(sim_expo_list_3$df_full, n_claim_2, prem_factor_3, desc=desc_3, claim_scale=0.05, line.idx=1, threshold_report=50000)

## Compare exposure to claims
# aggregate(sim_expo_list_1$df_full$Written_Premium,by=list(sim_expo_list_1$df_full$Policy_Year),sum)
# aggregate(sim_claim_list_1$df_claim$Ultimate_Loss,by=list(sim_claim_list_1$df_claim$Policy_Year),sum)
# aggregate(sim_expo_list_2$df_full$Written_Premium,by=list(sim_expo_list_2$df_full$Policy_Year),sum)
# aggregate(sim_claim_list_2$df_claim$Ultimate_Loss,by=list(sim_claim_list_2$df_claim$Policy_Year),sum)
# aggregate(sim_expo_list_3$df_full$Written_Premium,by=list(sim_expo_list_3$df_full$Policy_Year),sum)
# aggregate(sim_claim_list_3$df_claim$Ultimate_Loss,by=list(sim_claim_list_3$df_claim$Policy_Year),sum)

# WP by State & Segment
wp_state_new_1=data.frame(sim_expo_list_1$wp_state_new)
colnames(wp_state_new_1)=AYs_char[which(n_pol_1>0)]
wp_state_new_1$'2025'=round(apply(sim_expo_list_1$wp_state_new[,c("2023","2024")],1,mean)*(1+apply(rate_mat_1[,c("2023","2024")],1,mean)/100)^1.5,-3)
wp_state_new_2=data.frame(sim_expo_list_2$wp_state_new)
colnames(wp_state_new_2)=AYs_char[which(n_pol_2>0)]
wp_state_new_2$'2025'=round(apply(sim_expo_list_2$wp_state_new[,c("2023","2024")],1,mean)*(1+apply(rate_mat_2[,c("2023","2024")],1,mean)/100)^1.5,-3)
wp_state_new_3=data.frame(sim_expo_list_3$wp_state_new)
colnames(wp_state_new_3)=AYs_char[which(n_pol_3>0)]
wp_state_new_3$'2025'=round(apply(sim_expo_list_3$wp_state_new[,c("2023","2024")],1,mean)*(1+apply(rate_mat_3[,c("2023","2024")],1,mean)/100)^1.5,-3)
wp_state_old_1=data.frame(sim_expo_list_1$wp_state_old)
colnames(wp_state_old_1)=head(AYs_char[which(n_pol_1>0)],-1)
wp_state_old_1$'2024'=round(apply(sim_expo_list_1$wp_state_old[,c("2022","2023")],1,mean)*(1+apply(rate_mat_1[,c("2022","2023")],1,mean)/100)^1.5,-3)
wp_state_old_2=data.frame(sim_expo_list_2$wp_state_old)
colnames(wp_state_old_2)=head(AYs_char[which(n_pol_2>0)],-1)
wp_state_old_2$'2024'=round(apply(sim_expo_list_2$wp_state_old[,c("2022","2023")],1,mean)*(1+apply(rate_mat_2[,c("2022","2023")],1,mean)/100)^1.5,-3)
wp_state_old_3=data.frame(sim_expo_list_3$wp_state_old)
colnames(wp_state_old_3)=head(AYs_char[which(n_pol_3>0)],-1)
wp_state_old_3$'2024'=round(apply(sim_expo_list_3$wp_state_old[,c("2022","2023")],1,mean)*(1+apply(rate_mat_3[,c("2022","2023")],1,mean)/100)^1.5,-3)

# Claims
claim_list_new_1=data.frame(sim_claim_list_1$df_claim_output_new)
claim_list_new_2=data.frame(sim_claim_list_2$df_claim_output_new)
claim_list_new_3=data.frame(sim_claim_list_3$df_claim_output_new)
claim_list_old_1=data.frame(sim_claim_list_1$df_claim_output_old)
claim_list_old_2=data.frame(sim_claim_list_2$df_claim_output_old)
claim_list_old_3=data.frame(sim_claim_list_3$df_claim_output_old)

# rate change
rate_change_new_1=rate_change_1
rate_change_new_2=rate_change_2
rate_change_new_3=rate_change_3
rate_change_old_1=rate_change_1[rate_change_1$Year<year(LOSS_VAL_DATE),]
rate_change_old_2=rate_change_2[rate_change_2$Year<year(LOSS_VAL_DATE),]
rate_change_old_3=rate_change_3[rate_change_3$Year<year(LOSS_VAL_DATE),]
prem_factor_1
prem_factor_2
prem_factor_3

# Limit profile
sim_expo_list_1$limit_profile_new
sim_expo_list_2$limit_profile_new
sim_expo_list_3$limit_profile_new
sim_expo_list_1$limit_profile_old
sim_expo_list_2$limit_profile_old
sim_expo_list_3$limit_profile_old


####Paste Excel (new) ####
ps_text1 <-'
$ErrorActionPreference = "Stop"
$template = "%s"
$outfile  = "%s"
$csv      = "%s"
$sheet    = "'
ps_text2 <- '"
$cell     = "'
ps_text3 <- '"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($template)
$ws = $wb.Worksheets.Item($sheet)

$csvWb = $excel.Workbooks.Open($csv)
$csvWs = $csvWb.Worksheets.Item(1)

$used = $csvWs.UsedRange
$used.Copy() | Out-Null

# Paste values only (xlPasteValues = -4163)
$ws.Range($cell).PasteSpecial(-4163) | Out-Null

$csvWb.Close($false)
$wb.SaveAs($outfile)
$wb.Close($false)

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
'
template   <- "C:/Users/tfung/Desktop/Research/CAS grant AI/sample templates/broker submission/Sample2_UmbrellaXoL_empty.xlsx"
wb_out_new <- "C:/Users/tfung/Desktop/Research/CAS grant AI/sample templates/broker submission/Sample2_UmbrellaXoL_2025.xlsx"
## claim subline 1
csv_tmp <- tempfile(fileext = ".csv")
write.table( claim_list_new_1, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Large Claims (subline 1)",ps_text2,"A8",ps_text3), normalizePath(template, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## claim subline 2
csv_tmp <- tempfile(fileext = ".csv")
write.table( claim_list_new_2, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Large Claims (subline 2)",ps_text2,"A8",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## claim subline 3
csv_tmp <- tempfile(fileext = ".csv")
write.table( claim_list_new_3, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Large Claims (subline 3)",ps_text2,"A8",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## rate change subline 1
csv_tmp <- tempfile(fileext = ".csv")
write.table( rate_change_new_1, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Rate Change (subline 1)",ps_text2,"A9",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## rate change subline 2
csv_tmp <- tempfile(fileext = ".csv")
write.table( rate_change_new_2, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Rate Change (subline 2)",ps_text2,"A9",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## rate change subline 3
csv_tmp <- tempfile(fileext = ".csv")
write.table( rate_change_new_3, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Rate Change (subline 3)",ps_text2,"A9",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## wp subline 1
csv_tmp <- tempfile(fileext = ".csv")
write.table( wp_state_new_1, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"WP by State & Segment",ps_text2,"B64",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## wp subline 2
csv_tmp <- tempfile(fileext = ".csv")
write.table( wp_state_new_2, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"WP by State & Segment",ps_text2,"B119",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## wp subline 3
csv_tmp <- tempfile(fileext = ".csv")
write.table( wp_state_new_3, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"WP by State & Segment",ps_text2,"B174",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## limit profile subline 1
csv_tmp <- tempfile(fileext = ".csv")
write.table(sim_expo_list_1$limit_profile_new$x, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Limit Profile (in-force)",ps_text2,"B17",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## limit profile subline 2
csv_tmp <- tempfile(fileext = ".csv")
write.table(sim_expo_list_2$limit_profile_new$x, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Limit Profile (in-force)",ps_text2,"B26",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## limit profile subline 3
csv_tmp <- tempfile(fileext = ".csv")
write.table(sim_expo_list_3$limit_profile_new$x, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Limit Profile (in-force)",ps_text2,"B35",ps_text3), normalizePath(wb_out_new, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_new, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))



####Paste Excel (old) ####
ps_text1 <-'
$ErrorActionPreference = "Stop"
$template = "%s"
$outfile  = "%s"
$csv      = "%s"
$sheet    = "'
ps_text2 <- '"
$cell     = "'
ps_text3 <- '"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($template)
$ws = $wb.Worksheets.Item($sheet)

$csvWb = $excel.Workbooks.Open($csv)
$csvWs = $csvWb.Worksheets.Item(1)

$used = $csvWs.UsedRange
$used.Copy() | Out-Null

# Paste values only (xlPasteValues = -4163)
$ws.Range($cell).PasteSpecial(-4163) | Out-Null

$csvWb.Close($false)
$wb.SaveAs($outfile)
$wb.Close($false)

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
'
template   <- "C:/Users/tfung/Desktop/Research/CAS grant AI/sample templates/broker submission/Sample2_UmbrellaXoL_empty_old.xlsx"
wb_out_old <- "C:/Users/tfung/Desktop/Research/CAS grant AI/sample templates/broker submission/Sample2_UmbrellaXoL_2024.xlsx"
## claim subline 1
csv_tmp <- tempfile(fileext = ".csv")
write.table( claim_list_old_1, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Large Claims (subline 1)",ps_text2,"A8",ps_text3), normalizePath(template, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## claim subline 2
csv_tmp <- tempfile(fileext = ".csv")
write.table( claim_list_old_2, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Large Claims (subline 2)",ps_text2,"A8",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## claim subline 3
csv_tmp <- tempfile(fileext = ".csv")
write.table( claim_list_old_3, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Large Claims (subline 3)",ps_text2,"A8",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## rate change subline 1
csv_tmp <- tempfile(fileext = ".csv")
write.table( rate_change_old_1, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Rate Change (subline 1)",ps_text2,"A9",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## rate change subline 2
csv_tmp <- tempfile(fileext = ".csv")
write.table( rate_change_old_2, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Rate Change (subline 2)",ps_text2,"A9",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## rate change subline 3
csv_tmp <- tempfile(fileext = ".csv")
write.table( rate_change_old_3, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Rate Change (subline 3)",ps_text2,"A9",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## wp subline 1
csv_tmp <- tempfile(fileext = ".csv")
write.table( wp_state_old_1, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"WP by State & Segment",ps_text2,"B64",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## wp subline 2
csv_tmp <- tempfile(fileext = ".csv")
write.table( wp_state_old_2, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"WP by State & Segment",ps_text2,"B119",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## wp subline 3
csv_tmp <- tempfile(fileext = ".csv")
write.table( wp_state_old_3, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"WP by State & Segment",ps_text2,"B174",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## limit profile subline 1
csv_tmp <- tempfile(fileext = ".csv")
write.table(sim_expo_list_1$limit_profile_old$x, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Limit Profile (in-force)",ps_text2,"B17",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## limit profile subline 2
csv_tmp <- tempfile(fileext = ".csv")
write.table(sim_expo_list_2$limit_profile_old$x, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Limit Profile (in-force)",ps_text2,"B26",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))
## limit profile subline 3
csv_tmp <- tempfile(fileext = ".csv")
write.table(sim_expo_list_3$limit_profile_old$x, file  = csv_tmp, sep  = ",", row.names = FALSE, col.names = FALSE, quote = TRUE, na  = "")
ps1 <- tempfile(fileext = ".ps1")
ps_code <- sprintf(paste0(ps_text1,"Limit Profile (in-force)",ps_text2,"B35",ps_text3), normalizePath(wb_out_old, winslash="\\", mustWork=TRUE),
                   normalizePath(wb_out_old, winslash="\\", mustWork=FALSE),
                   normalizePath(csv_tmp, winslash="\\", mustWork=TRUE))
writeLines(ps_code, ps1)
system2("powershell",c("-NoProfile","-ExecutionPolicy","Bypass","-File", shQuote(ps1)))




# 
# ############ Output in CSV ##############
# write.csv(wp_state_new_1, file = "Sample2_XoL_wp_state_new_1.csv", row.names = T)
# write.csv(wp_state_new_2, file = "Sample2_XoL_wp_state_new_2.csv", row.names = T)
# write.csv(wp_state_new_3, file = "Sample2_XoL_wp_state_new_3.csv", row.names = T)
# write.csv(wp_state_old_1, file = "Sample2_XoL_wp_state_old_1.csv", row.names = T)
# write.csv(wp_state_old_2, file = "Sample2_XoL_wp_state_old_2.csv", row.names = T)
# write.csv(wp_state_old_3, file = "Sample2_XoL_wp_state_old_3.csv", row.names = T)
# write.csv(claim_list_new_1, file = "Sample2_XoL_claim_list_new_1.csv", row.names = F)
# write.csv(claim_list_new_2, file = "Sample2_XoL_claim_list_new_2.csv", row.names = F)
# write.csv(claim_list_new_3, file = "Sample2_XoL_claim_list_new_3.csv", row.names = F)
# write.csv(claim_list_old_1, file = "Sample2_XoL_claim_list_old_1.csv", row.names = F)
# write.csv(claim_list_old_2, file = "Sample2_XoL_claim_list_old_2.csv", row.names = F)
# write.csv(claim_list_old_3, file = "Sample2_XoL_claim_list_old_3.csv", row.names = F)
# write.csv(rate_change_new_1, file = "Sample2_XoL_rate_change_new_1.csv", row.names = F)
# write.csv(rate_change_new_2, file = "Sample2_XoL_rate_change_new_2.csv", row.names = F)
# write.csv(rate_change_new_3, file = "Sample2_XoL_rate_change_new_3.csv", row.names = F)
# write.csv(rate_change_old_1, file = "Sample2_XoL_rate_change_old_1.csv", row.names = F)
# write.csv(rate_change_old_2, file = "Sample2_XoL_rate_change_old_2.csv", row.names = F)
# write.csv(rate_change_old_3, file = "Sample2_XoL_rate_change_old_3.csv", row.names = F)




