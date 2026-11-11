rm(list = ls())

setwd("~/Report/Daily/Projection Report")




library(odbc)
library(DBI)
library(tidyverse)
library(dplyr)
library(openxlsx2)
library(gt)
library(lubridate)
library(bizdays)
library(mondate)
library(emayili)

con <- dbConnect(odbc:: odbc(),
                 Driver   = "*****",
                 Server   = "*****",
                 Database = "*****",
                 UID = "*****",
                 PWD = "*****",
                 Port = *****)

Validation  <- tbl(con, DBI::Id(schema = "SME", table = "Holiday")) %>% collect()
BD_Calender <- create.calendar(name = "BD", holidays = Validation$Date, weekdays = c("friday", "saturday"))

last_working_day  <- as.Date(bizdays::add.bizdays(Sys.Date(), -1, "BD")) #-1
prior_working_day <- as.Date(bizdays::add.bizdays(Sys.Date(), -2, "BD")) #-2
window_start_day  <- as.Date(last_working_day - months(3))

#NOTE: window is now anchored to Lead_Approver_Date (lead creation date), same 3-month
#trailing window used in the Rmd. Plain dates (no time component) are used, since
#Lead_Approver_Date is always cast to DATE in the queries below.
window_start <- paste0(window_start_day)
window_end   <- paste0(last_working_day)
yday         <- as.character(last_working_day)

#Current-month window -- used by the Sanctioned / Rejected / Insufficient Docs & Del.
#Nego. pulls below, same window as the Rmd's Final Decision summary table.
month_start_day <- as.Date(bizdays::adjust.next(lubridate::floor_date(last_working_day, "month"), "BD"))

excluded_branches     <- c("Uttara", "Dhanmondi")
excluded_branches_sql <- paste0("'", excluded_branches, "'", collapse = ", ")

excluded_rejection_reasons <- c(
  "Decision Pending for More Than 90 Days",
  "Sufficient Documents Not Submitted as Per the STP Policy",
  "Rejected Due to Delayed Negotiation"
)
excluded_rejection_sql <- paste0("'", excluded_rejection_reasons, "'", collapse = ", ")

#Branch -> RCM / CM mapping

branch_mapping <- tibble::tribble(
  ~Branch,                    ~RCM,                       ~CM,
  "Bogra",                    "Md. Kamal Uddin",          "Md. Kamal Uddin",
  "Dinajpur",                 "Md. Kamal Uddin",          "Md. Ziaur Rahman",
  "Kushtia",                  "Md. Kamal Uddin",          "Md. Mehedi Hasan",
  "Naogaon (Bogra)",          "Md. Kamal Uddin",          "Md. Kamal Uddin",
  "Natore",                   "Md. Kamal Uddin",          "Md. Mehedi Hasan",
  "Pabna (Natore)",           "Md. Kamal Uddin",          "Md. Mehedi Hasan",
  "Rajshahi",                 "Md. Kamal Uddin",          "Md. Mehedi Hasan",
  "Rangpur",                  "Md. Kamal Uddin",          "Md. Ziaur Rahman",
  "Barisal",                  "Tito Proshanta Rozario",   "Tito Proshanta Rozario",
  "Bhola (Barisal)",          "Tito Proshanta Rozario",   "Tito Proshanta Rozario",
  "Bhulta",                   "Tito Proshanta Rozario",   "Mahir Tazwer",
  "Elephant Road",            "Tito Proshanta Rozario",   "Mahir Tazwer",
  "Faridpur",                 "Tito Proshanta Rozario",   "Tito Proshanta Rozario",
  "Gulshan",                  "Tito Proshanta Rozario",   "Mahir Tazwer",
  "Jessore",                  "Tito Proshanta Rozario",   "Mohammad Shahadat Hossain",
  "Khulna",                   "Tito Proshanta Rozario",   "Mohammad Shahadat Hossain",
  "Mirpur",                   "Tito Proshanta Rozario",   "Md. Rakibul Alam",
  "Narsingdi",                "Tito Proshanta Rozario",   "Mahir Tazwer",
  "Savar",                    "Tito Proshanta Rozario",   "Md. Rakibul Alam",
  "Shatkhira (Jessore)",      "Tito Proshanta Rozario",   "Mohammad Shahadat Hossain",
  "Dilkusha",                 "S.M. Farabi Ferdous",      "Koushik Mohammad",
  "Gazipur",                  "S.M. Farabi Ferdous",      "Md. Rakibul Alam",
  "Imamganj",                 "S.M. Farabi Ferdous",      "Arman Ahmed",
  "Keranigonj",               "S.M. Farabi Ferdous",      "Arman Ahmed",
  "Mymensingh",               "S.M. Farabi Ferdous",      "Koushik Mohammad",
  "Narayangonj",              "S.M. Farabi Ferdous",      "Koushik Mohammad",
  "Sherpur (Mymensingh)",     "S.M. Farabi Ferdous",      "Koushik Mohammad",
  "Tangail (Gazipur)",        "S.M. Farabi Ferdous",      "Md. Rakibul Alam",
  "Tongi",                    "S.M. Farabi Ferdous",      "Koushik Mohammad",
  "Agrabad",                  "Mohammad Foysal Hasan",    "Mohammad Foysal Hasan",
  "Brahmanbaria (Comilla)",   "Mohammad Foysal Hasan",    "Mohammad Foysal Hasan",
  "Chowmuhani",               "Mohammad Foysal Hasan",    "Mohammad Foysal Hasan",
  "Comilla",                  "Mohammad Foysal Hasan",    "Mohammad Foysal Hasan",
  "Feni (Chowmuhani)",        "Mohammad Foysal Hasan",    "Mohammad Foysal Hasan",
  "Hobiganj",                 "Mohammad Foysal Hasan",    "Md. Raisul Azad",
  "Nandankanon",              "Mohammad Foysal Hasan",    "Mohammad Foysal Hasan",
  "Sylhet",                   "Mohammad Foysal Hasan",    "Md. Raisul Azad"
)

#RCM full name -> short nickname, used to label the RCM column in the dashboard
rcm_nickname_map <- c(
  "Md. Kamal Uddin"        = "KUddin",
  "Tito Proshanta Rozario" = "Tito",
  "S.M. Farabi Ferdous"    = "Ferdous",
  "Mohammad Foysal Hasan"  = "Foysal"
)

#RCM display order -- order of first appearance in branch_mapping
rcm_order <- unique(branch_mapping$RCM)

#Branch -> HoRB mapping, sourced from sme.RBH. Joined into every lead-wise detail tab
#(added after the STM column) -- not used in the summary/dashboard tables.
HoRB_mapping <- dbGetQuery(con, "SELECT Branch, RBH AS HoRB FROM sme.RBH")

#STP definition -- now sourced from dbo.AnalyticsOCASV2, matched to dbo.AnalyticsOCAS by
#LeadReferenceNo. FileCategory in AnalyticsOCASV2 ('STP' / otherwise) is used directly.
#A lead with no match in AnalyticsOCASV2 is left out entirely (INNER JOIN), so it never
#appears in any of the pulls below.
stp_join_sql <- "INNER JOIN dbo.AnalyticsOCASV2 ON dbo.AnalyticsOCAS.LeadReferenceNo = dbo.AnalyticsOCASV2.LeadReferenceNo"

#If FileCategory in AnalyticsOCASV2 is blank/NULL for a matched lead, treat it as Non-STP
#(rather than dropping it -- only a missing LeadReferenceNo match drops the lead, via the
#INNER JOIN above).
file_category_sql <- "CASE WHEN dbo.AnalyticsOCASV2.FileCategory = 'STP' THEN 'STP' ELSE 'Non-STP' END"


#To Be Visited
#NOTE: window is now anchored to Lead_Approver_Date (lead creation date), same 3-month
#trailing window used everywhere else. Prospect_Approve_Date_Latest is no longer
#date-range filtered -- it is only checked for non-null (the lead has reached this
#stage), so a lead created within the window still shows up here regardless of how
#recent/old its own stage date is.

to_be_visited_raw <- dbGetQuery(con,paste0("SELECT Branch,dbo.AnalyticsOCAS.LeadReferenceNo LeadReferenceNo,ProspectName,Interested_Amount_To,Interested_Amount_To/1E6 Interested_Amount_To_Mln,
[Business Type],CRG_Applicability,Product_Category,LOWER(Analyst_Name) Analyst_Name,[Credit Analyst] as Credit_Analyst,RM_Name,STM_Name,
CAST(Lead_Approver_Date AS DATE) Lead_Approver_Date,
CAST(Prospect_Approve_Date_Latest AS DATE) Prospect_Approve_Date_Latest,
", file_category_sql, " File_Category
FROM dbo.AnalyticsOCAS
", stp_join_sql, "
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND Product_Category = 'SME'
AND Lead_Approver_Date BETWEEN '", window_start, "' AND '", window_end, "'
AND Prospect_Approve_Date_Latest IS NOT NULL
AND CRM_Visit_Date IS NULL
AND Analyst_Pending_Collection_Date_Latest IS NULL
AND Analyst_Pending_Declaration_date_Initial IS NULL
AND Sanction_Date_Initial IS NULL
AND (Reject_Date IS NULL OR Reject_Date = '1900-01-01')")) |>
  mutate(
    Credit_Analyst = ifelse(is.na(Credit_Analyst) | trimws(Credit_Analyst) == "","Unassigned", Credit_Analyst),
    File_Category = ifelse(is.na(File_Category) | trimws(File_Category) == "","Non-STP", File_Category)) |>
  left_join(branch_mapping, by = "Branch") |>
  left_join(HoRB_mapping, by = "Branch")



to_be_visited_summary <- to_be_visited_raw |>
  group_by(Branch) |>
  summarise(`To Be Visited Number` = n(), `To Be Visited Amount` = sum(Interested_Amount_To_Mln)) |>
  as.data.frame() |>
  left_join(branch_mapping, by = "Branch")

to_be_visited_leads <- to_be_visited_raw |>
  transmute(
    `Lead Reference No` = LeadReferenceNo,
    Branch,
    Client               = ProspectName,
    `Proposed Amount`    = Interested_Amount_To,
    `File Category`      = File_Category,
    `CRG `       = CRG_Applicability,
    `Lead Creation Date` = Lead_Approver_Date,
    `Prospect Completion Date`    = Prospect_Approve_Date_Latest,
    Analyst              = Credit_Analyst,
    CM,
    RCM,
    RM                   = RM_Name,
    STM                  = STM_Name,
    HoRB
  )


#Pending To Be Declared
#Window anchored to Lead_Approver_Date; CRM_Visit_Date only checked for non-null.

pending_to_be_declared_raw <- dbGetQuery(con,paste0("SELECT Branch,dbo.AnalyticsOCAS.LeadReferenceNo LeadReferenceNo,ProspectName,Interested_Amount_To,Interested_Amount_To/1E6 Interested_Amount_To_Mln,
[Business Type],CRG_Applicability,LOWER(Analyst_Name) Analyst_Name,[Credit Analyst] as Credit_Analyst,RM_Name,STM_Name,
CAST(Lead_Approver_Date AS DATE) Lead_Approver_Date,
CAST(Prospect_Approve_Date_Latest AS DATE) Prospect_Approve_Date_Latest,
CAST(CRM_Visit_Date AS DATE) CRM_Visit_Date,
", file_category_sql, " File_Category
FROM dbo.AnalyticsOCAS
", stp_join_sql, "
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND Product_Category = 'SME'
AND ", file_category_sql, " = 'Non-STP'
AND Lead_Approver_Date BETWEEN '", window_start, "' AND '", window_end, "'
AND Prospect_Approve_Date_Latest IS NOT NULL
AND CRM_Visit_Date IS NOT NULL
AND Analyst_Pending_Collection_Date_Latest IS NULL
AND Analyst_Pending_Declaration_date_Initial IS NULL
AND Sanction_Date_Initial IS NULL
AND (Reject_Date IS NULL OR Reject_Date = '1900-01-01')")) |>
  mutate(
    Credit_Analyst = ifelse(is.na(Credit_Analyst) | trimws(Credit_Analyst) == "","Unassigned", Credit_Analyst),
    File_Category = ifelse(is.na(File_Category) | trimws(File_Category) == "","Non-STP", File_Category)) |>
  left_join(branch_mapping, by = "Branch") |>
  left_join(HoRB_mapping, by = "Branch")

pending_to_be_declared_summary <- pending_to_be_declared_raw |>
  group_by(Branch) |>
  summarise(`Pending To Be Declared Number` = n(), `Pending To Be Declared Amount` = sum(Interested_Amount_To_Mln)) |>
  as.data.frame() |>
  left_join(branch_mapping, by = "Branch")

pending_to_be_declared_leads <- pending_to_be_declared_raw |>
  transmute(
    `Lead Reference No` = LeadReferenceNo,
    Branch,
    Client               = ProspectName,
    `Proposed Amount`    = Interested_Amount_To,
    `File Category`      = File_Category,
    `CRG `       = CRG_Applicability,
    `Lead Creation Date` = Lead_Approver_Date,
    `Prospect Completion Date`    = Prospect_Approve_Date_Latest,
    `Analyst Visit Date` = CRM_Visit_Date,
    Analyst              = Credit_Analyst,
    CM,
    RCM,
    RM                   = RM_Name,
    STM                  = STM_Name,
    HoRB
  )


#Decision To Be Made -- STP leads pulled separately, then Non-STP leads, then stacked
#Window anchored to Lead_Approver_Date; every intermediate stage date is only checked for non-null.

decision_stp_leads <- dbGetQuery(con,paste0("SELECT Branch,dbo.AnalyticsOCAS.LeadReferenceNo LeadReferenceNo,ProspectName,Interested_Amount_To,Interested_Amount_To/1E6 Interested_Amount_To_Mln,
[Business Type],CRG_Applicability,LOWER(Analyst_Name) Analyst_Name,[Credit Analyst] as Credit_Analyst,RM_Name,STM_Name,
CAST(Lead_Approver_Date AS DATE) Lead_Approver_Date,
CAST(Prospect_Approve_Date_Latest AS DATE) Prospect_Approve_Date_Latest,
CAST(CRM_Visit_Date AS DATE) CRM_Visit_Date,
CAST(Analyst_Pending_Declaration_date_Initial AS DATE) Analyst_Pending_Declaration_date_Initial,
CAST(Analyst_Pending_Collection_Date_Latest AS DATE) Analyst_Pending_Collection_Date_Latest,
dbo.AnalyticsOCASV2.RequestDelayReason RequestDelayReason,
dbo.AnalyticsOCASV2.AnalystRemark AnalystRemark
FROM dbo.AnalyticsOCAS
", stp_join_sql, "
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND Product_Category = 'SME'
AND ", file_category_sql, " = 'STP'
AND Lead_Approver_Date BETWEEN '", window_start, "' AND '", window_end, "'
AND Prospect_Approve_Date_Latest IS NOT NULL
AND CRM_Visit_Date IS NOT NULL
AND Sanction_Date_Initial IS NULL
AND (Reject_Date IS NULL OR Reject_Date = '1900-01-01')")) |>
  mutate(File_Category  = "STP",
         Credit_Analyst = ifelse(is.na(Credit_Analyst) | trimws(Credit_Analyst) == "","Unassigned", Credit_Analyst)
  )

decision_non_stp_leads <- dbGetQuery(con,paste0("SELECT Branch,dbo.AnalyticsOCAS.LeadReferenceNo LeadReferenceNo,ProspectName,Interested_Amount_To,Interested_Amount_To/1E6 Interested_Amount_To_Mln,
[Business Type],CRG_Applicability,LOWER(Analyst_Name) Analyst_Name,[Credit Analyst] as Credit_Analyst,RM_Name,STM_Name,
CAST(Lead_Approver_Date AS DATE) Lead_Approver_Date,
CAST(Prospect_Approve_Date_Latest AS DATE) Prospect_Approve_Date_Latest,
CAST(CRM_Visit_Date AS DATE) CRM_Visit_Date,
CAST(Analyst_Pending_Declaration_date_Initial AS DATE) Analyst_Pending_Declaration_date_Initial,
CAST(Analyst_Pending_Collection_Date_Latest AS DATE) Analyst_Pending_Collection_Date_Latest,
dbo.AnalyticsOCASV2.RequestDelayReason RequestDelayReason,
dbo.AnalyticsOCASV2.AnalystRemark AnalystRemark
FROM dbo.AnalyticsOCAS
", stp_join_sql, "
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND Product_Category = 'SME'
AND ", file_category_sql, " = 'Non-STP'
AND Lead_Approver_Date BETWEEN '", window_start, "' AND '", window_end, "'
AND Prospect_Approve_Date_Latest IS NOT NULL
AND CRM_Visit_Date IS NOT NULL
AND Analyst_Pending_Collection_Date_Latest IS NOT NULL
AND Analyst_Pending_Declaration_date_Initial IS NOT NULL
AND Sanction_Date_Initial IS NULL
AND (Reject_Date IS NULL OR Reject_Date = '1900-01-01')")) |>
  mutate(File_Category  = "Non-STP",
         
         Credit_Analyst = ifelse(is.na(Credit_Analyst) | trimws(Credit_Analyst) == "","Unassigned", Credit_Analyst))

decision_to_be_made_raw <- rbind(decision_stp_leads, decision_non_stp_leads) |>
  left_join(branch_mapping, by = "Branch") |>
  left_join(HoRB_mapping, by = "Branch")

decision_to_be_made_summary <- decision_to_be_made_raw |>
  group_by(Branch) |>
  summarise(`Decision To Be Made Number` = n(), `Decision To Be Made Amount` = sum(Interested_Amount_To_Mln)) |>
  as.data.frame() |>
  left_join(branch_mapping, by = "Branch")

decision_to_be_made_leads <- decision_to_be_made_raw |>
  transmute(
    `Lead Reference No` = LeadReferenceNo,
    Branch,
    Client                     = ProspectName,
    `Proposed Amount`          = Interested_Amount_To,
    `File Category`            = File_Category,
    `CRG `       = CRG_Applicability,
    `Lead Creation Date`       = Lead_Approver_Date,
    `Prospect Completion Date`    = Prospect_Approve_Date_Latest,
    `Analyst Visit Date`       = CRM_Visit_Date,
    `Pending Declaration Date` = Analyst_Pending_Declaration_date_Initial,
    `Pending Collection Date`  = Analyst_Pending_Collection_Date_Latest,
    `Delayed Processing Reason`  = RequestDelayReason,
    `Delayed Processing Remarks` = AnalystRemark,
    Analyst                    = Credit_Analyst,
    CM,
    RCM,
    RM                         = RM_Name,
    STM                        = STM_Name,
    HoRB
  )


#Delayed Processing -- same leads as "Decision To Be Made" above (same STP/Non-STP pulls,
#same stage logic), further restricted to leads that haven't moved since last_working_day:
#for STP leads, Analyst Visit Date (CRM_Visit_Date) is not last_working_day; for Non-STP
#leads, Pending Collection Date (Analyst_Pending_Collection_Date_Latest) is not
#last_working_day. Same column set/order as the "Decision To Be Made" tab.

delayed_processing_raw <- decision_to_be_made_raw |>
  filter(
    (File_Category == "STP"     & CRM_Visit_Date != last_working_day) |
      (File_Category == "Non-STP" & Analyst_Pending_Collection_Date_Latest != last_working_day)
  )

delayed_processing_leads <- delayed_processing_raw |>
  transmute(
    `Lead Reference No` = LeadReferenceNo,
    Branch,
    Client                     = ProspectName,
    `Proposed Amount`          = Interested_Amount_To,
    `File Category`            = File_Category,
    `CRG `       = CRG_Applicability,
    `Lead Creation Date`       = Lead_Approver_Date,
    `Prospect Completion Date`    = Prospect_Approve_Date_Latest,
    `Analyst Visit Date`       = CRM_Visit_Date,
    `Pending Declaration Date` = Analyst_Pending_Declaration_date_Initial,
    `Pending Collection Date`  = Analyst_Pending_Collection_Date_Latest,
    `Delayed Processing Reason`  = RequestDelayReason,
    `Delayed Processing Remarks` = AnalystRemark,
    Analyst                    = Credit_Analyst,
    CM,
    RCM,
    RM                         = RM_Name,
    STM                        = STM_Name,
    HoRB
  )


#Pending Declared But Not Collected -- Non-STP leads that have been declared but not yet
#collected (Pending Collection Date is still NULL), not yet sanctioned or rejected.
#Window anchored to Lead_Approver_Date; Analyst_Pending_Declaration_date_Initial only
#checked for non-null.

pending_declared_not_collected_raw <- dbGetQuery(con,paste0("SELECT Branch,dbo.AnalyticsOCAS.LeadReferenceNo LeadReferenceNo,ProspectName,Interested_Amount_To,Interested_Amount_To/1E6 Interested_Amount_To_Mln,
[Business Type],CRG_Applicability,LOWER(Analyst_Name) Analyst_Name,[Credit Analyst] as Credit_Analyst,RM_Name,STM_Name,
CAST(Lead_Approver_Date AS DATE) Lead_Approver_Date,
CAST(Prospect_Approve_Date_Latest AS DATE) Prospect_Approve_Date_Latest,
CAST(CRM_Visit_Date AS DATE) CRM_Visit_Date,
CAST(Analyst_Pending_Declaration_date_Initial AS DATE) Analyst_Pending_Declaration_date_Initial, Analyst_Pending_Remark,
", file_category_sql, " File_Category
FROM dbo.AnalyticsOCAS
", stp_join_sql, "
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND Product_Category = 'SME'
AND ", file_category_sql, " = 'Non-STP'
AND Lead_Approver_Date BETWEEN '", window_start, "' AND '", window_end, "'
AND Prospect_Approve_Date_Latest IS NOT NULL
AND CRM_Visit_Date IS NOT NULL
AND Analyst_Pending_Declaration_date_Initial IS NOT NULL
AND Analyst_Pending_Collection_Date_Latest IS NULL
AND Sanction_Date_Initial IS NULL
AND (Reject_Date IS NULL OR Reject_Date = '1900-01-01')")) |>
  mutate(
    Credit_Analyst = ifelse(is.na(Credit_Analyst) | trimws(Credit_Analyst) == "","Unassigned", Credit_Analyst),
    File_Category = ifelse(is.na(File_Category) | trimws(File_Category) == "","Non-STP", File_Category)
  ) |>
  left_join(branch_mapping, by = "Branch") |>
  left_join(HoRB_mapping, by = "Branch")

pending_declared_not_collected_leads <- pending_declared_not_collected_raw |>
  transmute(
    `Lead Reference No` = LeadReferenceNo,
    Branch,
    Client                     = ProspectName,
    `Proposed Amount`          = Interested_Amount_To,
    `File Category`            = File_Category,
    `CRG `                     = CRG_Applicability,
    `Lead Creation Date`       = Lead_Approver_Date,
    `Prospect Completion Date` = Prospect_Approve_Date_Latest,
    `Analyst Visit Date`       = CRM_Visit_Date,
    `Pending Declaration Date` = Analyst_Pending_Declaration_date_Initial,
    Analyst                    = Credit_Analyst,
    CM,
    RCM,
    RM                         = RM_Name,
    STM                        = STM_Name,
    HoRB,
    `Pending Remarks`          = Analyst_Pending_Remark
  )


#Sanctioned / Rejected / Insufficient Docs & Del. Nego. -- lead-wise detail tabs, current
#month window (month_start_day through last_working_day), matching the Rmd's Final
#Decision summary table window.

#-- Sanctioned: Final Approver is resolved from sme.Approver (Appraisal_Approver_Final
#is a userid/coe code in AnalyticsOCAS; sme.Approver.userid maps it to a full name in
#the "Approver Name" column).
sanctioned_raw <- dbGetQuery(con,paste0("SELECT Branch,dbo.AnalyticsOCAS.LeadReferenceNo LeadReferenceNo,ProspectName,Interested_Amount_To,Sanc_Amt,
CRG_Applicability,[Credit Analyst] as Credit_Analyst,RM_Name,STM_Name,
CAST(Lead_Approver_Date AS DATE) Lead_Approver_Date,
CAST(Prospect_Approve_Date_Latest AS DATE) Prospect_Approve_Date_Latest,
CAST(CRM_Visit_Date AS DATE) CRM_Visit_Date,
CAST(Analyst_Pending_Declaration_date_Initial AS DATE) Analyst_Pending_Declaration_date_Initial,
CAST(Analyst_Pending_Collection_Date_Latest AS DATE) Analyst_Pending_Collection_Date_Latest,
CAST(Sanction_Date_Initial AS DATE) Sanction_Date_Initial,
ap.[Approver Name] Final_Approver,
", file_category_sql, " File_Category
FROM dbo.AnalyticsOCAS
", stp_join_sql, "
LEFT JOIN sme.Approver ap ON dbo.AnalyticsOCAS.Appraisal_Approver_Final = ap.userid
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND Product_Category = 'SME'
AND Sanction_Date_Initial IS NOT NULL
AND CAST(Sanction_Date_Initial AS DATE) BETWEEN '", month_start_day, "' AND '", last_working_day, "'")) |>
  mutate(
    Credit_Analyst = ifelse(is.na(Credit_Analyst) | trimws(Credit_Analyst) == "","Unassigned", Credit_Analyst),
    File_Category = ifelse(is.na(File_Category) | trimws(File_Category) == "","Non-STP", File_Category)) |>
  left_join(branch_mapping, by = "Branch") |>
  left_join(HoRB_mapping, by = "Branch")

sanctioned_leads <- sanctioned_raw |>
  transmute(
    `Lead Reference No`        = LeadReferenceNo,
    Branch,
    Client                     = ProspectName,
    `Proposed Amount`          = Interested_Amount_To,
    `Sanctioned Amount`        = Sanc_Amt,
    `File Category`            = File_Category,
    `CRG `                     = CRG_Applicability,
    `Lead Creation Date`       = Lead_Approver_Date,
    `Prospect Completion Date` = Prospect_Approve_Date_Latest,
    `Analyst Visit Date`       = CRM_Visit_Date,
    `Pending Declaration Date` = Analyst_Pending_Declaration_date_Initial,
    `Pending Collection Date`  = Analyst_Pending_Collection_Date_Latest,
    `Sanction Date`            = Sanction_Date_Initial,
    Analyst                    = Credit_Analyst,
    `Final Approver`           = Final_Approver,
    CM,
    RCM,
    RM                         = RM_Name,
    STM                        = STM_Name,
    HoRB
  )

#-- Rejected + Insufficient Docs & Del. Nego. -- pulled together, then split by
#Reason_Of_Rejection into the two tabs.
rejected_all_raw <- dbGetQuery(con,paste0("SELECT Branch,dbo.AnalyticsOCAS.LeadReferenceNo LeadReferenceNo,ProspectName,Interested_Amount_To,
CRG_Applicability,[Credit Analyst] as Credit_Analyst,RM_Name,STM_Name,
CAST(Lead_Approver_Date AS DATE) Lead_Approver_Date,
CAST(Prospect_Approve_Date_Latest AS DATE) Prospect_Approve_Date_Latest,
CAST(CRM_Visit_Date AS DATE) CRM_Visit_Date,
CAST(Analyst_Pending_Declaration_date_Initial AS DATE) Analyst_Pending_Declaration_date_Initial,
CAST(Analyst_Pending_Collection_Date_Latest AS DATE) Analyst_Pending_Collection_Date_Latest,
CAST(Reject_Date AS DATE) Reject_Date,
Reason_Of_Rejection,
dbo.AnalyticsOCASV2.RejectReasonDetail RejectReasonDetail,
", file_category_sql, " File_Category
FROM dbo.AnalyticsOCAS
", stp_join_sql, "
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND Product_Category = 'SME'
AND Reject_Date IS NOT NULL
AND Reject_Date <> '1900-01-01'
AND CAST(Reject_Date AS DATE) BETWEEN '", month_start_day, "' AND '", last_working_day, "'")) |>
  mutate(
    Credit_Analyst = ifelse(is.na(Credit_Analyst) | trimws(Credit_Analyst) == "","Unassigned", Credit_Analyst),
    File_Category = ifelse(is.na(File_Category) | trimws(File_Category) == "","Non-STP", File_Category)) |>
  left_join(branch_mapping, by = "Branch") |>
  left_join(HoRB_mapping, by = "Branch")

rejected_raw     <- rejected_all_raw |> filter(!Reason_Of_Rejection %in% excluded_rejection_reasons)
insufficient_raw <- rejected_all_raw |> filter(Reason_Of_Rejection %in% excluded_rejection_reasons)

rejected_leads <- rejected_raw |>
  transmute(
    `Lead Reference No`        = LeadReferenceNo,
    Branch,
    Client                     = ProspectName,
    `Proposed Amount`          = Interested_Amount_To,
    `File Category`            = File_Category,
    `CRG `                     = CRG_Applicability,
    `Lead Creation Date`       = Lead_Approver_Date,
    `Prospect Completion Date` = Prospect_Approve_Date_Latest,
    `Analyst Visit Date`       = CRM_Visit_Date,
    `Pending Declaration Date` = Analyst_Pending_Declaration_date_Initial,
    `Pending Collection Date`  = Analyst_Pending_Collection_Date_Latest,
    `Rejection Date`           = Reject_Date,
    `Rejection Reason`         = Reason_Of_Rejection,
    `Rejection Remarks`        = RejectReasonDetail,
    Analyst                    = Credit_Analyst,
    CM,
    RCM,
    RM                         = RM_Name,
    STM                        = STM_Name,
    HoRB
  )

insufficient_leads <- insufficient_raw |>
  transmute(
    `Lead Reference No`        = LeadReferenceNo,
    Branch,
    Client                     = ProspectName,
    `Proposed Amount`          = Interested_Amount_To,
    `File Category`            = File_Category,
    `CRG `                     = CRG_Applicability,
    `Lead Creation Date`       = Lead_Approver_Date,
    `Prospect Completion Date` = Prospect_Approve_Date_Latest,
    `Analyst Visit Date`       = CRM_Visit_Date,
    `Pending Declaration Date` = Analyst_Pending_Declaration_date_Initial,
    `Pending Collection Date`  = Analyst_Pending_Collection_Date_Latest,
    `Rejection Date`           = Reject_Date,
    `Rejection Reason`         = Reason_Of_Rejection,
    `Rejection Remarks`        = RejectReasonDetail,
    Analyst                    = Credit_Analyst,
    CM,
    RCM,
    RM                         = RM_Name,
    STM                        = STM_Name,
    HoRB
  )


#Waiting for Prospect Creation -- leads whose Lead Creation Date (Lead_Approver_Date, date
#only, no time component) falls in the trailing-90-day window (window_start_day through
#last_working_day), but which have not yet reached Prospect Approve stage.
#(Already anchored to Lead_Approver_Date -- unchanged.)
#NOTE: this tab was previously named "Waiting for Lead Creation" -- renamed to "Waiting
#for Prospect Creation". File Category and CRG are now pulled in as well (via the same
#STP join/logic used by every other tab), since this tab didn't carry them before.

waiting_for_lead_creation_raw <- dbGetQuery(con,paste0("SELECT Branch,dbo.AnalyticsOCAS.LeadReferenceNo LeadReferenceNo,ProspectName,Interested_Amount_To,Interested_Amount_To/1E6 Interested_Amount_To_Mln,
[Business Type],CRG_Applicability,RM_Name,STM_Name,
CAST(Lead_Approver_Date AS DATE) Lead_Approver_Date,
", file_category_sql, " File_Category
FROM dbo.AnalyticsOCAS
", stp_join_sql, "
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND Product_Category = 'SME'
AND CAST(Lead_Approver_Date AS DATE) BETWEEN '", window_start_day, "' AND '", last_working_day, "'
AND Prospect_Approve_Date_Latest IS NULL")) |>
  mutate(File_Category = ifelse(is.na(File_Category) | trimws(File_Category) == "","Non-STP", File_Category)) |>
  left_join(branch_mapping, by = "Branch") |>
  left_join(HoRB_mapping, by = "Branch")

waiting_for_lead_creation_leads <- waiting_for_lead_creation_raw |>
  transmute(
    LeadReferenceNo,
    Branch,
    Client                = ProspectName,
    `Proposed Amount`     = Interested_Amount_To,
    `File Category`       = File_Category,
    `CRG `       = CRG_Applicability,
    `Lead Creation Date`  = Lead_Approver_Date,
    CM,
    RCM,
    RM                    = RM_Name,
    STM                   = STM_Name,
    HoRB
  )


#Yesterday's Status -- Visited / Pending Declared / Decision Finalized (exact date match on last_working_day)
#Yesterday's Status -- Visited / Pending Declared / Decision Finalized (exact date match on last_working_day)
visited_summary <- dbGetQuery(con,paste0("SELECT Branch, COUNT(*) [Visited Number], SUM(Interested_Amount_To)/1E6 [Visited Amount]
FROM dbo.AnalyticsOCAS
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND CAST(CRM_Visit_Date AS DATE) = '", yday, "'
AND Product_Category = 'SME'
GROUP BY Branch")) |>
  left_join(branch_mapping, by = "Branch")

pending_declared_summary <- dbGetQuery(con,paste0("SELECT Branch, COUNT(*) [Pending Declared Number], SUM(Interested_Amount_To)/1E6 [Pending Declared Amount]
FROM dbo.AnalyticsOCAS
INNER JOIN dbo.AnalyticsOCASV2 ON dbo.AnalyticsOCAS.LeadReferenceNo = dbo.AnalyticsOCASV2.LeadReferenceNo
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND ", file_category_sql, " = 'Non-STP'
AND CAST(Analyst_Pending_Declaration_date_Initial AS DATE) = '", yday, "'
AND Product_Category = 'SME'
GROUP BY Branch")) |>
  left_join(branch_mapping, by = "Branch")

decision_finalized_sanctioned <- dbGetQuery(con,paste0("SELECT Branch, LeadReferenceNo, Sanc_Amt/1E6 Interested_Amount_To
FROM dbo.AnalyticsOCAS
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND CAST(Sanction_Date_Initial AS DATE) = '", yday, "'
AND Product_Category = 'SME'"))

decision_finalized_rejected <- dbGetQuery(con,paste0("SELECT Branch, LeadReferenceNo, Interested_Amount_To/1E6 Interested_Amount_To
FROM dbo.AnalyticsOCAS
WHERE Branch NOT IN (", excluded_branches_sql, ")
AND CAST(Reject_Date AS DATE) = '", yday, "'
AND Reason_Of_Rejection NOT IN (", excluded_rejection_sql, ")
AND Product_Category = 'SME'"))

decision_finalized_summary <- rbind(decision_finalized_sanctioned, decision_finalized_rejected) |>
  group_by(Branch) |>
  summarise(`Decision Finalized Number` = n(), `Decision Finalized Amount` = sum(Interested_Amount_To)) |>
  as.data.frame() |>
  left_join(branch_mapping, by = "Branch")


#Sync filter -- gate on whether yesterday's saved projection file (needed for the
#Dashboard's "Yesterday's Projection" comparison) is present; if it's missing the
#Dashboard would show zeros for that block, so route to self only in that case.

output_dir     <- "~/Report/Daily/Projection Report"
yesterday_file <- file.path(output_dir, paste0("Projection_for_", prior_working_day, ".xlsx"))


#Send_to_me <- !file.exists(yesterday_file)

# sync

df <- tbl(con, DBI::Id(schema = "dbo", table = "AnalyticsOCAS")) %>% 
  filter(Product_Category == "SME")


df1old <- tbl(con, DBI::Id(schema = "dbo", table = "AnalyticsCLAccount")) %>% 
  filter(PRODUCT_SUBCATEGORY == "SEF")


base_tbl <- tbl(con, DBI::Id(schema = "dbo", table = "AnalyticsCL_Disbursement")) 

df1 <- left_join(base_tbl %>% select(ACCOUNT_NUMBER, LeadReferenceNo, TOTAL_AMOUNT, VALUE_DT, EXECUTION_DATE, MAKER_DT_STAMP,
                                     DSBR_STATUS, MAKER_ID, CHECKER_ID) %>% 
                   rename(AMOUNT_DISBURSED = TOTAL_AMOUNT,
                          LAST_DISB_DATE = VALUE_DT) 
                 ,df1old %>% select(ACCOUNT_NUMBER, ALT_ACC_NO,SETTLEMENT_ACC,CUSTOMER_NO, CLIENTNAME:TERMINMONTH,ACCOUNT_STATUS:Mobile_Number, SectorGroup:RMName),
                 by = "ACCOUNT_NUMBER") %>% 
  filter(DSBR_STATUS == "P", PRODUCT_SUBCATEGORY == "SEF")



Validation <- tbl(con, DBI::Id(schema = "SME", table = "Holiday")) %>% collect()

BD_Calender <- create.calendar(name = "BD", holidays = Validation$Date, weekdays = c("friday", "saturday"))

smtp <- server(host = "192.168.110.5", port = 26)

last_working_day <- as.Date(bizdays::add.bizdays(Sys.Date(), -1, "BD"))

#last_working_day <- as.Date("2024-12-31")


df1 %>% mutate(LAST_DISB_DATE = as.Date(LAST_DISB_DATE)) %>%
  filter(ACCOUNT_NUMBER == ALT_ACC_NO, LAST_DISB_DATE == last_working_day) %>% 
  count(LAST_DISB_DATE) %>% 
  collect() %>% 
  pull(n) -> last_day_disb

df %>% 
  mutate(Sanction_Date_Initial = as.Date(Sanction_Date_Initial)) %>% 
  select(Sanction_Date_Initial) %>% 
  filter(Sanction_Date_Initial == last_working_day) %>% 
  count(Sanction_Date_Initial) %>% 
  collect() %>% 
  pull(n) -> last_day_sanc


##"How many records have OCAS disbursement date equal to last_working_day but Flexcube initial disbursement date is missing."

df |> 
  mutate(Date_of_Disbursement_Execution_Flexcube_Initial=as.Date(Date_of_Disbursement_Execution_Flexcube_Initial),
         Disbursement_Execution_in_OCAS=as.Date(Disbursement_Execution_in_OCAS)) |> 
  filter(Disbursement_Execution_in_OCAS==last_working_day) |> 
  count(p=is.na(Date_of_Disbursement_Execution_Flexcube_Initial)) |> 
  filter(p==TRUE) |> 
  summarise(n=n()) |> 
  pull(n) -> last_day_mismatch

Send_to_me <- !(length(last_day_sanc) > 0 & length(last_day_disb) > 0 & !last_day_mismatch>0)
#Send_to_me <- TRUE


#Yesterday's Projection -- read back the plan saved in yesterday's workbook.
#NOTE: the saved Dashboard sheet is now RCM-wise (col A = RCM, col B = Branch, data
#starts row 4, and includes "Sub-total"/"Total" rows interspersed) -- those rows are
#dropped here so only real branch rows feed into yesterday_record.

today_cols <- c("To Be Visited Number", "To Be Visited Amount",
                "Pending To Be Declared Number", "Pending To Be Declared Amount",
                "Decision To Be Made Number", "Decision To Be Made Amount")

if (file.exists(yesterday_file)) {
  raw <- openxlsx2::wb_load(yesterday_file) |>
    openxlsx2::wb_to_df(sheet = "Dashboard", start_row = 4, col_names = FALSE)
  names(raw)[1] <- "RCM_Display"
  names(raw)[2] <- "Branch"
  names(raw)[3:8] <- today_cols
  yesterday_record <- raw |>
    filter(!Branch %in% c("Sub-total", "Sub Total", "Total")) |>
    select(Branch, all_of(today_cols)) |>
    mutate(across(all_of(today_cols), as.numeric))
} else {
  yesterday_record <- data.frame(Branch = character(), check.names = FALSE)
  yesterday_record[today_cols] <- numeric()
}

yesterday_record <- yesterday_record |>
  rename(
    `Yesterday_To Be Visited Number`          = `To Be Visited Number`,
    `Yesterday_To Be Visited Amount`          = `To Be Visited Amount`,
    `Yesterday_Pending To Be Declared Number` = `Pending To Be Declared Number`,
    `Yesterday_Pending To Be Declared Amount` = `Pending To Be Declared Amount`,
    `Yesterday_Decision To Be Made Number`    = `Decision To Be Made Number`,
    `Yesterday_Decision To Be Made Amount`    = `Decision To Be Made Amount`
  )


#Dashboard -- join every piece by Branch, same list + reduce(left_join) pattern as the daily report
#branch_mapping now also contributes RCM so the dashboard can be re-cut RCM-wise below.

list_df <- list(
  branch_mapping |> select(Branch, RCM),
  to_be_visited_summary          |> select(Branch, `To Be Visited Number`, `To Be Visited Amount`),
  pending_to_be_declared_summary |> select(Branch, `Pending To Be Declared Number`, `Pending To Be Declared Amount`),
  decision_to_be_made_summary    |> select(Branch, `Decision To Be Made Number`, `Decision To Be Made Amount`),
  yesterday_record,
  visited_summary                |> select(Branch, `Visited Number`, `Visited Amount`),
  pending_declared_summary       |> select(Branch, `Pending Declared Number`, `Pending Declared Amount`),
  decision_finalized_summary     |> select(Branch, `Decision Finalized Number`, `Decision Finalized Amount`)
)

dashboard_df <- list_df |> reduce(left_join, by = "Branch")

dashboard_df[is.na(dashboard_df)] <- 0

dashboard_df <- dashboard_df |> filter(!Branch %in% excluded_branches)


#RCM-wise re-cut of the dashboard.
#Each RCM's branches are grouped together (RCM label shown once per group, blank for the
#remaining branch rows of that RCM so it reads as a merged cell), followed by a Sub-total
#row for that RCM, with a final grand Total row at the bottom.

metric_cols <- c(
  "To Be Visited Number", "To Be Visited Amount",
  "Pending To Be Declared Number", "Pending To Be Declared Amount",
  "Decision To Be Made Number", "Decision To Be Made Amount",
  "Yesterday_To Be Visited Number", "Yesterday_To Be Visited Amount",
  "Visited Number", "Visited Amount",
  "Yesterday_Pending To Be Declared Number", "Yesterday_Pending To Be Declared Amount",
  "Pending Declared Number", "Pending Declared Amount",
  "Yesterday_Decision To Be Made Number", "Yesterday_Decision To Be Made Amount",
  "Decision Finalized Number", "Decision Finalized Amount"
)

build_rcm_dashboard <- function(df, metric_cols, rcm_order, nickname_map) {
  
  df <- df |>
    mutate(RCM = factor(RCM, levels = rcm_order)) |>
    arrange(RCM, Branch)
  
  group_blocks <- lapply(rcm_order, function(rcm_name) {
    grp <- df |> filter(RCM == rcm_name)
    if (nrow(grp) == 0) return(NULL)
    
    branch_rows <- grp |>
      mutate(
        RCM_Display = c(unname(nickname_map[rcm_name]), rep(NA_character_, dplyr::n() - 1)),
        Row_Type    = "Branch"
      ) |>
      select(RCM_Display, Branch, Row_Type, all_of(metric_cols)) |>
      mutate(Group = rcm_name)
    
    subtotal_row <- grp |>
      summarise(across(all_of(metric_cols), sum)) |>
      mutate(RCM_Display = NA_character_, Branch = "Sub-total", Row_Type = "Subtotal", Group = rcm_name) |>
      select(RCM_Display, Branch, Row_Type, Group, all_of(metric_cols))
    
    bind_rows(branch_rows, subtotal_row)
  })
  
  result <- bind_rows(group_blocks)
  
  total_row <- df |>
    summarise(across(all_of(metric_cols), sum)) |>
    mutate(RCM_Display = NA_character_, Branch = "Total", Row_Type = "Total", Group = NA_character_) |>
    select(RCM_Display, Branch, Row_Type, Group, all_of(metric_cols))
  
  bind_rows(result, total_row)
}

rcm_dashboard_df <- build_rcm_dashboard(dashboard_df, metric_cols, rcm_order, rcm_nickname_map)


out <- wb_workbook()

#Dashboard -- placed first so it appears before "To Be Visited"; RCM-wise with a merged
#RCM column per group, a Sub-total row after each RCM's branches, and a grand Total row.
out$add_worksheet("Dashboard", grid_lines = FALSE, zoom = 90)

out$add_data(sheet = "Dashboard", x = "RCM",    start_row = 1, start_col = 1, col_names = FALSE)
out$add_data(sheet = "Dashboard", x = "Branch", start_row = 1, start_col = 2, col_names = FALSE)
out$add_data(sheet = "Dashboard", x = "Today's Projection",     start_row = 1, start_col = 3, col_names = FALSE)
out$add_data(sheet = "Dashboard", x = "Yesterday's Projection", start_row = 1, start_col = 9, col_names = FALSE)
out$merge_cells(sheet = "Dashboard", dims = "A1:A3")
out$merge_cells(sheet = "Dashboard", dims = "B1:B3")
out$merge_cells(sheet = "Dashboard", dims = "C1:H1")
out$merge_cells(sheet = "Dashboard", dims = "I1:T1")

spanner_labels     <- c("To Be Visited", "Pending To Be Declared", "Decision To Be Made",
                        "To Be Visited", "Visited", "Pending To Be Declared", "Pending Declared",
                        "Decision To Be Made", "Decision Finalized")
spanner_start_cols <- seq(3, 19, by = 2)

for (i in seq_along(spanner_labels)) {
  scol <- spanner_start_cols[i]
  out$add_data(sheet = "Dashboard", x = spanner_labels[i], start_row = 2, start_col = scol, col_names = FALSE)
  out$merge_cells(sheet = "Dashboard", dims = wb_dims(rows = 2, cols = scol:(scol + 1)))
}

out$add_data(sheet = "Dashboard", x = as.data.frame(t(rep(c("Number", "Amount(Mln)"), 9))),
             start_row = 3, start_col = 3, col_names = FALSE)

dashboard_body <- rcm_dashboard_df |> select(RCM_Display, Branch, all_of(metric_cols))
dashboard_body$RCM_Display[is.na(dashboard_body$RCM_Display)] <- ""

first_data_row <- 4
out$add_data(sheet = "Dashboard", x = dashboard_body, start_row = first_data_row, col_names = FALSE)

last_row <- first_data_row + nrow(dashboard_body) - 1

#-- Merge the RCM column down each RCM's branch rows (Sub-total / Total rows excluded)
branch_idx <- which(rcm_dashboard_df$Row_Type == "Branch")
run_group  <- rle(rcm_dashboard_df$Group[branch_idx])
pos <- 1
for (k in seq_along(run_group$lengths)) {
  len <- run_group$lengths[k]
  excel_rows <- branch_idx[pos:(pos + len - 1)] + first_data_row - 1
  if (len > 1) {
    out$merge_cells(sheet = "Dashboard", dims = wb_dims(rows = min(excel_rows):max(excel_rows), cols = 1))
  }
  pos <- pos + len
}

#-- Colours: RCM / Branch / Today's Projection header block = #F54927, Yesterday's = #F87C63
out$add_fill(sheet = "Dashboard", dims = wb_dims(rows = 1:3, cols = 1:8),  color = wb_color("#F54927"))
out$add_fill(sheet = "Dashboard", dims = wb_dims(rows = 1:3, cols = 9:20), color = wb_color("#F87C63"))
out$add_font(sheet = "Dashboard", dims = wb_dims(rows = 1:3, cols = 1:20), color = wb_color("#FFFFFF"),
             bold = TRUE, name = "Cambria", size = 10)

subtotal_excel_rows <- which(rcm_dashboard_df$Row_Type == "Subtotal") + first_data_row - 1
total_excel_rows    <- which(rcm_dashboard_df$Row_Type == "Total")    + first_data_row - 1

#wb_dims() only accepts contiguous row ranges, and Sub-total rows are scattered between
#RCM groups, so each row is styled individually here rather than in one wb_dims() call.
for (r in subtotal_excel_rows) {
  out$add_fill(sheet = "Dashboard", dims = wb_dims(rows = r, cols = 1:20), color = wb_color("#E0E0E0"))
  out$add_font(sheet = "Dashboard", dims = wb_dims(rows = r, cols = 1:20), bold = TRUE, name = "Cambria", size = 10)
}

for (r in total_excel_rows) {
  out$add_fill(sheet = "Dashboard", dims = wb_dims(rows = r, cols = 1:20), color = wb_color("#CCCCCC"))
  out$add_font(sheet = "Dashboard", dims = wb_dims(rows = r, cols = 1:20), bold = TRUE, name = "Cambria", size = 10)
}

out$add_border(sheet = "Dashboard", dims = wb_dims(rows = 1:last_row, cols = 1:20), bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")
out$add_font(sheet = "Dashboard", dims = wb_dims(rows = 4:last_row, cols = 1:20), name = "Cambria", size = 10)
out$set_col_widths(sheet = "Dashboard", cols = 1:20, widths = "auto")


out$add_worksheet("To Be Visited",grid_lines = FALSE, zoom = 90)
out$add_data(x = to_be_visited_leads,sheet = "To Be Visited")
out$add_border(sheet = "To Be Visited", dims = "A1:N1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")


out$add_font(
  sheet = current_sheet(),
  dims = "A1:N1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "To Be Visited", cols = 1:14, widths = "auto")
out$add_fill(sheet = "To Be Visited",dims = "A1:N1",color = wb_color("#FDFD96"))



out$add_worksheet("Pending To Be Declared",grid_lines = FALSE, zoom = 90)
out$add_data(x = pending_to_be_declared_leads,sheet = "Pending To Be Declared")
out$add_border(sheet = "Pending To Be Declared", dims = "A1:O1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")


out$add_font(
  sheet = current_sheet(),
  dims = "A1:O1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "Pending To Be Declared", cols = 1:15, widths = "auto")
out$add_fill(sheet = "Pending To Be Declared",dims = "A1:O1",color = wb_color("#FDFD96"))



out$add_worksheet("Decision To Be Made",grid_lines = FALSE, zoom = 90)
out$add_data(x = decision_to_be_made_leads,sheet = "Decision To Be Made")
out$add_border(sheet = "Decision To Be Made", dims = "A1:S1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")


out$add_font(
  sheet = current_sheet(),
  dims = "A1:S1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "Decision To Be Made", cols = 1:19, widths = "auto")

out$add_fill(sheet = "Decision To Be Made",dims = "A1:S1",color = wb_color("#FDFD96"))


#-- NEW TAB: Delayed Processing ----------------------------------------------------------
out$add_worksheet("Delayed Processing",grid_lines = FALSE, zoom = 90)
out$add_data(x = delayed_processing_leads,sheet = "Delayed Processing")
out$add_border(sheet = "Delayed Processing", dims = "A1:S1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")

out$add_font(
  sheet = current_sheet(),
  dims = "A1:S1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "Delayed Processing", cols = 1:19, widths = "auto")
out$add_fill(sheet = "Delayed Processing",dims = "A1:S1",color = wb_color("#FDFD96"))


#-- NEW TAB: Pending Declared But Not Collected -----------------------------------------
out$add_worksheet("Pending Declared Not Collected",grid_lines = FALSE, zoom = 90)
out$add_data(x = pending_declared_not_collected_leads,sheet = "Pending Declared Not Collected")
out$add_border(sheet = "Pending Declared Not Collected", dims = "A1:Q1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")

out$add_font(
  sheet = current_sheet(),
  dims = "A1:Q1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "Pending Declared Not Collected", cols = 1:17, widths = "auto")
out$add_fill(sheet = "Pending Declared Not Collected",dims = "A1:Q1",color = wb_color("#FDFD96"))


#-- NEW TAB: Sanctioned -------------------------------------------------------------------
out$add_worksheet("Sanctioned",grid_lines = FALSE, zoom = 90)
out$add_data(x = sanctioned_leads,sheet = "Sanctioned")
out$add_border(sheet = "Sanctioned", dims = "A1:T1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")

out$add_font(
  sheet = current_sheet(),
  dims = "A1:T1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "Sanctioned", cols = 1:20, widths = "auto")
out$add_fill(sheet = "Sanctioned",dims = "A1:T1",color = wb_color("#FDFD96"))


#-- NEW TAB: Rejected ---------------------------------------------------------------------
out$add_worksheet("Rejected",grid_lines = FALSE, zoom = 90)
out$add_data(x = rejected_leads,sheet = "Rejected")
out$add_border(sheet = "Rejected", dims = "A1:T1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")

out$add_font(
  sheet = current_sheet(),
  dims = "A1:T1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "Rejected", cols = 1:20, widths = "auto")
out$add_fill(sheet = "Rejected",dims = "A1:T1",color = wb_color("#FDFD96"))


#-- NEW TAB: Insufficient Docs & Del. Nego. (short name -- Excel's 31-char sheet-name
#limit doesn't fit "Insufficient Documents and Delayed Negotiation") -------------------
out$add_worksheet("Insufficient Docs & Del. Nego.",grid_lines = FALSE, zoom = 90)
out$add_data(x = insufficient_leads,sheet = "Insufficient Docs & Del. Nego.")
out$add_border(sheet = "Insufficient Docs & Del. Nego.", dims = "A1:T1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")

out$add_font(
  sheet = current_sheet(),
  dims = "A1:T1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "Insufficient Docs & Del. Nego.", cols = 1:20, widths = "auto")
out$add_fill(sheet = "Insufficient Docs & Del. Nego.",dims = "A1:T1",color = wb_color("#FDFD96"))


#-- NEW TAB: Waiting for Prospect Creation (formerly "Waiting for Lead Creation") -------
out$add_worksheet("Waiting for Prospect Creation",grid_lines = FALSE, zoom = 90)
out$add_data(x = waiting_for_lead_creation_leads,sheet = "Waiting for Prospect Creation")
out$add_border(sheet = "Waiting for Prospect Creation", dims = "A1:L1000", bottom_border = "thin",
               left_border = "thin",
               right_border = "thin",
               top_border = "thin",
               inner_hgrid = "thin",
               inner_vgrid = "thin")

out$add_font(
  sheet = current_sheet(),
  dims = "A1:L1000",
  name = "Cambria",
  size = "10")
out$set_col_widths(sheet = "Waiting for Prospect Creation", cols = 1:12, widths = "auto")
out$add_fill(sheet = "Waiting for Prospect Creation",dims = "A1:L1",color = wb_color("#FDFD96"))




date_subject <- format(as.Date(bizdays::add.bizdays(Sys.Date(), 0, "BD")), "%B %d, %Y") ### 0 chilo

excel_path <- paste0("~/Report/Daily/Projection Report/Todays_Projection_Report_Data_", 
                     format(as.Date(bizdays::add.bizdays(Sys.Date(), 0, "BD")), "%Y-%m-%d"), 
                     ".xlsx")

out$save(excel_path)

smtp <- emayili::server(host = "192.168.110.5", port = 26)

email_projection_self <- emayili::envelope() %>%
  emayili::from(addr = emayili::address(email = "CreditAnalytics@IDLC.com", display = "IDLC - Credit Analytics")) %>%
  emayili::to(c("Sunjida@IDLC.com", "Monayer@IDLC.com")) %>%
  emayili::subject(glue::glue("Daily Projection Report {date_subject}")) %>%
  emayili::render(input = "Todays_Projection.Rmd", include_css = c("rmd", "bootstrap","highlight")) |>
  emayili::attachment(excel_path)

email_projection <- emayili::envelope() %>%
  emayili::from(addr = emayili::address(email = "CreditAnalytics@IDLC.com", display = "IDLC - Credit Analytics")) %>%
  # emayili::to(c("Sunjida@IDLC.com", "Monayer@IDLC.com")) %>%
  emayili::to(c( "GrpCECSEF@IDLC.com", "Adnan@IDLC.com", "Areful@IDLC.com", "MMomin@IDLC.com",
                 "Samiul@IDLC.com")) %>%
  emayili::cc(c( "GrpSMELT@IDLC.com", "GrpCrdSBF@IDLC.com", "GrpSEFinCharge@IDLC.com", "GrpSMESTM@IDLC.com",
                 "Sunjida@IDLC.com", "Monayer@IDLC.com")) |>
  emayili::subject(glue::glue("Daily Projection Report {date_subject}")) %>%
  emayili::render(input = "Todays_Projection.Rmd", include_css = c("rmd", "bootstrap","highlight")) |>
  emayili::attachment(excel_path)

if(Send_to_me & bizdays::is.bizday(Sys.Date(), "BD")){
  smtp(email_projection_self, verbose = TRUE)}

if(!Send_to_me & bizdays::is.bizday(Sys.Date(), "BD")){
  smtp(email_projection, verbose = TRUE)}