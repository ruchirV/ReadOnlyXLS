//
//  ViewController.m
//  MyExcel
//
//  Created by Ruchir on 05/05/14.
//

#import "ViewController.h"
#import "UIExcelSheetView.h"
#import "ExcelSheetColumnData.h"

@interface ViewController ()

@end

@implementation ViewController

- (id)initWithNibName:(NSString *)nibNameOrNil bundle:(NSBundle *)nibBundleOrNil
{
    self = [super initWithNibName:nibNameOrNil bundle:nibBundleOrNil];
    
    if (self)
    {
        // Custom initialization
    }
    
    return self;
}

- (void)viewDidLoad
{
    [super viewDidLoad];
    
    [self.navigationItem setTitle:@"Main"];
}

- (IBAction)xlsButtonHandler:(UIButton *)sender
{
    static NSString* nameKey = @"name";
    static NSString* numVisitsKey = @"numberOfVisits";
    static NSString* emailId = @"emailId";
    
    NSMutableArray* dataArray = [NSMutableArray new];

    for (int iter = 1; iter <= 50; iter++)
    {
        NSString* name = [NSString stringWithFormat:@"Customer %d",iter];
        NSString* email = [NSString stringWithFormat:@"%@@gmail.com", [name lowercaseString]];
        
        [dataArray addObject:@{nameKey: name, numVisitsKey:@(iter), emailId:email}];
    }
    
    NSMutableArray* columnArray = [NSMutableArray new];
    
    ExcelSheetColumnData* colData = [ExcelSheetColumnData new];
    [colData setColumnTitle:@"Name"];
    [colData setPropertyKey:nameKey];
    [colData setWidth:200];
    [columnArray addObject:colData];
    
    colData = [ExcelSheetColumnData new];
    [colData setColumnTitle:@"Number of Visit"];
    [colData setPropertyKey:numVisitsKey];
    [colData setWidth:200];
    [columnArray addObject:colData];
    
    colData = [ExcelSheetColumnData new];
    [colData setColumnTitle:@"Email"];
    [colData setPropertyKey:emailId];
    [colData setWidth:300];
    [columnArray addObject:colData];
    
    [columnArray insertObject:[ExcelSheetColumnData IndexColumn] atIndex:0];
    
    UIViewController* xlsController = [[UIViewController alloc] init];
    [self.navigationController pushViewController:xlsController animated:NO];
    
    ExcelSheetTotalRowData* totalVisits = [ExcelSheetTotalRowData new];
    totalVisits.property = numVisitsKey;
    totalVisits.totalType = ExcelSheetTotalData_Sum;
    
    ExcelSheetTotalRowData* totalCustomers = [ExcelSheetTotalRowData new];
    totalCustomers.property = nameKey;
    totalCustomers.totalType = ExcelSheetTotalData_Count;
    
    UIExcelSheetView* xlsView = [[UIExcelSheetView alloc] initWithFrame:CGRectMake(0, 0, self.view.frame.size.width, self.view.frame.size.height)
                                                                   Data:dataArray
                                                                Columns:columnArray
                                                          TotalProperty:@[totalCustomers, totalVisits]
                                                           HeaderHeight:35.0
                                                              RowHeight:25.0
                                                        WithDefaultSort:nameKey
                                                              SortOrder:UIExcelSortOrderDescending];
    
    [xlsController.view addSubview:xlsView];
}

@end
