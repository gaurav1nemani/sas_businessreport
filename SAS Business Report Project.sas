/*SAS GROUP PROJECT - GROUP 4*/
/*Group Members: Gaurav Nemani, Paul Guillard, Mario Perez and Benjamin Ro*/

/*Define Library*/
LIBNAME Group4 'C:\Users\Source\OneDrive - IESEG\Documents\SEMESTER 1 COURSES\Business Analytics Tool - Commercial SAS\Group\Final Coding GN';

/*-------------------------------------------------------#######  PART I   #########-------------------------------------------------------*/

/*Read all the three datasets and sort them*/
DATA Group4.Customer_Table_missing;
	INFILE 'C:\Users\Source\OneDrive - IESEG\Documents\SEMESTER 1 COURSES\Business Analytics Tool - Commercial SAS\Group\Data group project\CustomerTable.dat' FIRSTOBS = 2 DSD DLM = ';';
	INPUT CustomerID Industry :$20. Region :$20. Revenue Employees;
	FORMAT Revenue DOLLAR25.2;
RUN;

DATA Group4.Order_Table;
	INFILE 'C:\Users\Source\OneDrive - IESEG\Documents\SEMESTER 1 COURSES\Business Analytics Tool - Commercial SAS\Group\Data group project\OrderTable.csv' FIRSTOBS = 2 DSD DLM=',';
	INPUT CustomerID OrderID OrderAmount Orderdate :MMDDYY8.;
	FORMAT OrderDate WORDDATE18.; 	/*Formatted Date for readability*/
	FORMAT OrderAmount DOLLAR25.2;
RUN;

PROC SORT DATA=Group4.Order_Table;
	BY CustomerId;
RUN;

DATA Group4.ESG_scores_CSV_missing;
	INFILE 'C:\Users\Source\OneDrive - IESEG\Documents\SEMESTER 1 COURSES\Business Analytics Tool - Commercial SAS\Group\Data group project\ESG_scores.csv' FIRSTOBS = 2 DLM=',' DSD;
	INPUT Industry :$14. Environmental Social Governance;
RUN;

PROC SORT DATA=Group4.ESG_scores_CSV_missing;
	BY Industry;
RUN;

DATA Group4.Manager_Table; 
	INPUT Account_Manager_Name $22. Industries $23-79;
	DATALINES; 
	Tim Tom           Technology, Finance, Healthcare 
	Jonathan Lee Wang Education, Energy, Manufacturing, Retail, Transportation 
	; 
RUN;

/*Making Manager Table more data readable*/
DATA Group4.Manager_Table;
	SET Group4.Manager_Table;
	LENGTH Responsible_for $14.;
	DO i=1 TO COUNTW(Industries, ', ');
		Responsible_for=SCAN(Industries, i, ', ');
		OUTPUT;
	END;
	DROP i Industries;
RUN;
PROC SORT DATA=Group4.Manager_Table;
	BY Responsible_for;
RUN;

/*-------------------------------------------------------#######  PART II   #########-------------------------------------------------------*/

ODS PDF FILE='C:\Users\Source\OneDrive - IESEG\Documents\SEMESTER 1 COURSES\Business Analytics Tool - Commercial SAS\Group\Final Coding GN\Report\Test\ExecutiveReportTest.pdf';

/*Showcase Missing Values in the executive summary*/

PROC MEANS DATA=Group4.Customer_Table_missing NMISS;
	VAR Revenue Employees;
	OUTPUT OUT = summary_missing_executive;
	TITLE 'Filling the Gaps: Tackling Missing Revenue and Employee Data';
RUN;

PROC FREQ DATA=Group4.Customer_Table_missing;
	TABLE Industry;
	TITLE 'Filling the Gaps: Tackling Missing Industry Data';
RUN;

/*Create dummy variables for missing values*/
DATA Group4.Customer_Table_missing1;
	SET Group4.Customer_Table_missing;
	dummy_Industry_missing= MISSING(Industry);
	dummy_Region_missing= MISSING(Region);
	dummy_Revenue_missing= MISSING(Revenue);
	dummy_Employees_missing= MISSING(Employees);
RUN;

/*Impute missing for numeric variables Revenue and Employees*/
PROC MEANS NOPRINT DATA=Group4.Customer_Table_missing1;
	VAR Revenue Employees;
	OUTPUT OUT=Group4.summarymeans 
	MEAN(Revenue Employees)= mean_Revenue mean_Employees
	SUM(Revenue)=Total_Revenue_sum;
RUN;

DATA Group4.Customer_Table_missing1;
	IF _N_=1 THEN SET Group4.summarymeans;
	SET Group4.Customer_Table_missing1;
	IF MISSING(Revenue) THEN Revenue= mean_Revenue;
	IF MISSING(Employees) THEN Employees= mean_Employees;
	IF MISSING(Account_Manager_Name) THEN Account_Manager_Name='Tim Tom';
	DROP _TYPE_ _FREQ_ mean_Revenue mean_Employees Total_Revenue_sum;
RUN;

/*Impute missing for character variable Industry Using Hot Deck Imputation Method {Simple Random Selection With Replacement} */
PROC SURVEYIMPUTE NOPRINT METHOD=HOTDECK (SELECTION=SRSWR) SEED=123;
	VAR Industry;
	OUTPUT OUT=Group4.Cust_HotDeck;
RUN;

DATA Group4.Customer_Table;
	MERGE Group4.Customer_Table_missing1 Group4.Cust_HotDeck;
	BY CustomerID;
RUN;
PROC SORT DATA=Group4.Customer_Table;
	BY CustomerID;
RUN;

/*Create Aggregations*/
PROC MEANS NOPRINT DATA=Group4.Order_Table MAXDEC=2;
	BY CustomerID;
	VAR OrderAmount OrderId Orderdate;
	OUTPUT OUT=Group4.Summaryagg
	MIN(Orderdate)=First_OrderDate
	MAX(Orderdate)=Last_OrderDate
	;
RUN;
	
DATA Group4.Order_Table_Agg;
	MERGE Group4.Order_Table Group4.Summaryagg;
	BY CustomerID;
	Recency_days=TODAY()-Last_OrderDate;
	DROP _TYPE_ _FREQ_;
RUN;
PROC SORT DATA=Group4.Order_Table_Agg;
	BY CustomerID;
RUN;

DATA Group4.Esg_scores_csv_missing;
	SET Group4.Esg_scores_csv_missing;
	Mean_ESG_Score=(Environmental+Social+Governance)/3;
RUN;

/*Join into final BaseTable and add the aggregate values: Avg_Order, Recency*/
PROC SQL; 
	CREATE TABLE Group4.BaseTable AS 
	SELECT ct.CustomerID, ct.Industry, ct.Region, ct.Revenue, ct.Employees, ct.dummy_Industry_missing, ct.dummy_Region_missing, ct.dummy_Revenue_missing, ct.dummy_Employees_missing, 
			COUNT(ot.OrderId) AS Nbr_Orders, SUM(ot.OrderAmount) AS Total_OrderAmount, AVG(ot.OrderAmount) AS Avg_Order, MIN(ot.OrderDate) AS First_OrderDate, MAX(ot.OrderDate) AS Last_OrderDate, MAX(ot.Recency_days) AS Recency_days, 
			esg.Environmental, esg.Social, esg.Governance, esg.Mean_ESG_Score,
			mt.Account_Manager_Name
	FROM Group4.Customer_Table AS ct LEFT JOIN Group4.Order_Table_Agg AS ot ON ct.CustomerID = ot.CustomerID 
			LEFT JOIN Group4.ESG_scores_CSV_missing AS esg ON ct.Industry = esg.Industry 
			LEFT JOIN Group4.Manager_Table AS mt ON ct.Industry = mt.Responsible_for 
	GROUP BY 1,2,3,4,5,6,7,8,9,16,17,18,19, 20;
QUIT;

/*EXTRA AGGREGATIONS DONE:
1. Mean of ESG Scores
2. Avg Order Frequency(Avg Purchase Frequency)
3. Assumed that Customer Retention rate is 80% so Average Customer Lifespan is (1/(1-0.8))=5 years 
4. Customer Lifetime Value = (Avg Orders* Avg Order Frequency)*Avg Customer Lifespan
*/

DATA Group4.BaseTable;
	SET Group4.BaseTable;
	FORMAT First_OrderDate Last_OrderDate WORDDATE18.
	Avg_Order Total_OrderAmount DOLLAR18.;
RUN;

PROC MEANS NOPRINT DATA=Group4.BaseTable N;
	OUTPUT OUT = Group4.summary_clv N(CustomerID)=Nbr_customers;
RUN;

DATA Group4.BaseTable;
	IF _N_=1 THEN SET Group4.summary_clv;
	SET Group4.BaseTable;
	Avg_Order_Freq=Nbr_Orders/Nbr_customers;
	Customer_Lifetime_Value= (Avg_Order*Avg_Order_Freq)*5;
	DROP _TYPE_ _FREQ_ Nbr_Customers;
RUN;

PROC PRINT DATA=Group4.BaseTable (OBS=5) NOOBS;
	TITLE 'Quick Peek: First 5 Rows of the Customer Table';
RUN;

/*########################  PART - III   #######################*/

/*Insight 1: Customer Industry Distribution*/

PROC FREQ NOPRINT DATA=Group4.BaseTable;
	TABLES Industry / OUT=Group4.Summary_Insight1 (RENAME= (COUNT=Nbr_Customers));
RUN;

PROC SORT DATA = Group4.Summary_Insight1;
	BY DESCENDING Nbr_Customers;
RUN;

DATA Group4.Insight_1_CustIndusDist;
	SET Group4.Summary_Insight1;
	DROP Percent;
RUN;

PROC FORMAT;
	VALUE Nbr_Customers_color 0 -<100 = 'LIGHT RED'
						101 -< 115 = 'LIGHT YELLOW'
						116 -< 250 = 'LIGHT GREEN';
RUN;

PROC PRINT DATA=Group4.Insight_1_CustIndusDist NOOBS;
	VAR Industry;
	VAR Nbr_Customers / STYLE = {BACKGROUNDCOLOR= Nbr_Customers_color.};
	TITLE "Customer Distribution Across Industries";
RUN;

/*Insight 2: Avg (Number of) Orders per customer per industry*/

PROC MEANS NOPRINT DATA=Group4.Basetable;
	CLASS Industry;
	OUTPUT OUT=Group4.summary_Insight2 
	SUM(Nbr_Orders)=GrandTotal_Orders
	N(CustomerID)=Nbr_Customers;
RUN;

DATA Group4.Insight_2_AvgNOrders_CustIndus;
	SET Group4.summary_Insight2;
	AvgOrder_Cust_Ind=GrandTotal_Orders/Nbr_Customers;
	FORMAT AvgOrder_Cust_Ind 3.2;
	DROP _TYPE_ _FREQ_;
	IF MISSING(Industry) THEN DELETE;
RUN;

PROC FORMAT;
	VALUE GTO_color 338 -< 375 = 'LIGHT RED'
						375 -< 425 = 'LIGHT YELLOW'
						425 -< 600 = 'LIGHT GREEN';
	VALUE NC_color 99 -< 106 = 'LIGHT RED'
						106 -< 120 = 'LIGHT YELLOW'
						120 -< 600 = 'LIGHT GREEN';
	VALUE ACI_color 3.2 -< 3.4 = 'LIGHT RED'
						3.4 -< 3.6 = 'LIGHT YELLOW'
						3.6 -< 5 = 'LIGHT GREEN';
RUN;

PROC PRINT DATA=Group4.Insight_2_AvgNOrders_CustIndus NOOBS;
	VAR Industry;
	VAR GrandTotal_Orders / STYLE = {BACKGROUNDCOLOR= GTO_color.};
	VAR Nbr_Customers / STYLE = {BACKGROUNDCOLOR= NC_color.};
	VAR AvgOrder_Cust_Ind / STYLE = {BACKGROUNDCOLOR= ACI_color.};
	TITLE 'Industry Insights: Customer Counts Revealed';
RUN;

/*Insights 3 - Returning Customers*/

DATA CustomerClassification;
    SET Group4.Basetable(KEEP=CustomerID Nbr_Orders);
    
    IF Nbr_Orders > 1 THEN CustomerType = "Return Buyer";
    ELSE CustomerType = "First Buyer";
RUN;

PROC MEANS DATA=CustomerClassification NOPRINT;
    CLASS CustomerType;
    OUTPUT OUT=Group4.Insight_3_ReturnCust (DROP=_TYPE_ _FREQ_) N=Count;
RUN;

PROC GCHART DATA=Group4.Insight_3_ReturnCust;
    TITLE "Mapping the Loyalty: Returning Customer Trends"; 
    PIE CustomerType / SUMVAR=Count
                      PERCENT=INSIDE
                      VALUE=INSIDE
                      SLICE=OUTSIDE;
RUN;

/*Insights 4 - Number of Customers per Region*/

DATA RegionalSales;
    SET Group4.Basetable(KEEP=CustomerID Region);
RUN;

PROC SORT DATA=RegionalSales NODUPKEY OUT=UniqueCustomers;
    BY Region CustomerID;
RUN;

PROC MEANS DATA=UniqueCustomers NOPRINT;
    CLASS Region;
    OUTPUT OUT=RegionalSalesSummary (DROP=_TYPE_ _FREQ_)
        N(CustomerID)=Total_Customers;
RUN;
DATA Group4.Insight_4_NCusRegion;
    SET RegionalSalesSummary;
    IF missing(Region) THEN DELETE; 
RUN;

PROC SGPLOT DATA=Group4.Insight_4_NCusRegion;
    VBAR Region / RESPONSE=Total_Customers STAT=SUM DATALABEL;
    XAXIS LABEL="Region";
    YAXIS LABEL="Number of Customers";
    TITLE "Regional Breakdown: Customer Counts Unveiled";
RUN;

/*Insights 5 - Trend of Sales over time*/

DATA Group4.select_year;
	SET Group4.Order_table_agg;
	Order_Year=YEAR(OrderDate);
RUN;

DATA Group4.sales_with_month;
    SET Group4.Order_Table_Agg;
    month_order = month(Orderdate);
RUN;

PROC SQL;
    CREATE TABLE Group4.sales_by_month AS
    SELECT month_order, 
           SUM(OrderAmount) AS total_sales
    FROM Group4.sales_with_month
    GROUP BY month_order
    ORDER BY month_order;
QUIT;

PROC MEANS NOPRINT DATA= Group4.sales_with_month;
    CLASS month_order;
    VAR OrderAmount;
    OUTPUT OUT=Group4.sales_by_month_sum SUM=total_sales;
RUN;

PROC REG DATA=Group4.sales_by_month OUTEST=trendline_coeff NOPRINT;
    MODEL total_sales = month_order;
QUIT;

DATA Group4.insight_5_SalesTrend;
    SET Group4.sales_by_month;
    trend_sales = 620000 + 2000 * month_order; 
RUN;

PROC SGPLOT data=Group4.insight_5_SalesTrend;
    SERIES X=month_order Y=total_sales / MARKERS; 
    SERIES X=month_order Y=trend_sales / LINEATTRS=(COLOR=red PATTERN=shortdash THICKNESS=2); /* Trendline */
    XAXIS LABEL="Month of the Year" VALUES=(1 to 12 by 1)
          VALUESDISPLAY=('JAN' 'FEB' 'MAR' 'APR' 'MAY' 'JUN'
                         'JUL' 'AUG' 'SEP' 'OCT' 'NOV' 'DEC');
    YAXIS LABEL="Total Sales ($)" MIN=475000 MAX=800000; /* Limit defined manually */
    TITLE "Rising Success: Sales Trends Over All The Years";
RUN;

/*Insights 6 - Correlation between ESG to Total Order Amount*/

DATA Group4.Insight_6_CorrESG_OrderAmt;
	SET Group4.BaseTable;
RUN;

PROC SGSCATTER DATA=Group4.Insight_6_CorrESG_OrderAmt;
    MATRIX Environmental Social Governance Total_OrderAmount / 
        DIAGONAL=(HISTOGRAM) 
        MARKERATTRS=(SYMBOL=CIRCLEFILLED COLOR=BLUE);
    TITLE "Linking ESG to Revenue: Correlation with Total Order Amount";
    LABEL Environmental="Environmental (E)"
          Social="Social (S)"
          Governance="Governance (G)"
          Total_OrderAmount="Total Order Amount";
RUN;

PROC ODSTEXT;
	p "Exploring the Impact of ESG Metrics on Revenue Growth"/ style=[fontsize=11 just=center];
	p "";
	p "Vertical Clustering in Total Order Amount Scatterplots: " / style=[fontsize=10 just=left];
	p "The vertical alignment of points suggests that ESG scores have limited variability, causing multiple companies" / style=[fontsize=8 just=left]; 
	p "to have different Total Order Amounts despite having the same ESG scores." / style=[fontsize=8 just=left]; 
	p "";
	p "Social (S) Scores Are Concentrated in the Lower Range: " / style=[fontsize=10 just=left];
	p "The histogram for Social scores reveals that most companies perform relatively poorly in Social aspects compared to" / style=[fontsize=8 just=left];
	p "Environmental and Governance scores." / style=[fontsize=8 just=left];
	p "";
	p "Governance (G) Scores Skew Towards Higher Values: " / style=[fontsize=10 just=left];
	p "Governance scores despite being evenly distributed but showcase a slight skew towards higher values, which" / style=[fontsize=8 just=left]; 
	p "indicates better overall performance in governance compared to the other ESG components." / style=[fontsize=8 just=left];
	p "";
	p "Weak Positive Correlation Between Social (S) and Governance (G): " / style=[fontsize=10 just=left];
	p "Companies scoring higher in Governance tend to have moderately higher Social scores, depicting some " / style=[fontsize=8 just=left];
	p "interdependency between these dimensions." / style=[fontsize=8 just=left];
RUN;

/*########################  PART - IV   #######################*/

DATA Group4.Jonathan_Top_Cust Group4.Tim_Top_Cust;
	SET Group4.Basetable;
	IF Account_Manager_Name='Tim Tom' THEN OUTPUT Group4.Tim_Top_Cust;
	ELSE IF Account_Manager_Name='Jonathan Lee Wang' THEN OUTPUT Group4.Jonathan_Top_Cust;
RUN;

PROC SORT DATA=Group4.Jonathan_Top_Cust;
	BY DESCENDING Total_OrderAmount;
RUN;

PROC SORT DATA=Group4.Tim_Top_Cust;
	BY DESCENDING Total_OrderAmount;
RUN;


DATA Group4.Jonathan_Top_Cust;
	SET Group4.Jonathan_Top_Cust(OBS=20);
	KEEP CustomerID Industry Region Nbr_Orders Total_OrderAmount Avg_Order Recency_days Mean_ESG_Score Avg_Order_Freq Customr_Lifetime_Value;
RUN;

PROC PRINT DATA=Group4.Jonathan_Top_Cust NOOBS;
	TITLE 'Jonathan Lee Wang’s Star Clients: Top Customers Spotlight';
RUN;

DATA Group4.Tim_Top_Cust;
	SET Group4.Tim_Top_Cust (OBS=20);
	KEEP CustomerID Industry Region Nbr_Orders Total_OrderAmount Avg_Order Recency_days Mean_ESG_Score Avg_Order_Freq Customr_Lifetime_Value;
RUN;

PROC PRINT DATA=Group4.Tim_Top_Cust NOOBS;
	TITLE 'Tim Tom’s Star Clients: Top Customers Spotlight';
RUN;

/*########################  PART - V   #######################*/

%MACRO Yearly_sales_report (Year_ofReport=);
	
	DATA Group4.select_year;
		SET Group4.select_year;
		WHERE Order_Year=&Year_ofReport;
	RUN;

	PROC MEANS NOPRINT DATA=Group4.select_year MAXDEC=2 SUM;
		VAR OrderAmount;
		OUTPUT OUT=yearly_revenue SUM(OrderAmount)=Total_Yearly_Revenue;
	RUN;

	DATA Group4.yearly_report;
		SET yearly_revenue;
		Year_of_Report=&Year_ofReport;
		DROP _TYPE_ _FREQ_;
	RUN;

	PROC PRINT DATA=Group4.yearly_report NOOBS;
		VAR Total_Yearly_Revenue;
		FORMAT Total_Yearly_Revenue DOLLAR20.2;
		TITLE "&Year_ofReport Revenue Breakdown: Total Earnings Report";
	RUN;

%MEND Yearly_sales_report;

%Yearly_sales_report(Year_ofReport=2022)

DATA Group4.select_year;
	SET Group4.Order_table_agg;
	Order_Year=YEAR(OrderDate);
RUN;

ODS PDF CLOSE;
