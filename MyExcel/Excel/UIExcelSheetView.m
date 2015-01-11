//
//  UIExcelSheetView.m
//  MyExcel
//
//  Created by Ruchir on 12/05/14.
//

#import "UIExcelSheetView.h"
#import "ExcelSheetColumnData.h"
#import <QuartzCore/QuartzCore.h>

#define HEADER_HEIGHT 70
#define ROW_HEIGHT 65
#define LEFT_PADDING 4
#define RIGHT_PADDING 4

#define __LOG__     //NSLog(@"%s", __FUNCTION__);


#define COLUMNSTART_TAG 11

#define CELL_LABEL_TAG 1
#define CELL_BUTTON_TAG 3

#define DEFAULT_HEADER_FONTSIZE 19.0f
#define DEFAULT_CELL_FONTSIZE 17.0f

@interface UIExcelSheetView()

@property (nonatomic, strong) UIScrollView                  *tableScrollView;
@property (nonatomic, strong) NSMutableArray                *xlsData;
@property (nonatomic, strong) NSArray                       *columnData;
@property (nonatomic, strong) UITableView                   *tableView;
@property (nonatomic, strong) NSString                      *sortProperty;
@property (nonatomic, assign) UIExcelSortOrder              sortOrder;
@property (nonatomic, strong) NSLayoutConstraint            *tableHeightConstraint;
@property (nonatomic, assign) CGFloat                       headerHeight;
@property (nonatomic, assign) CGFloat                       rowHeight;
@property (nonatomic, strong) NSArray*                      totalRowData;

@end

@implementation UIExcelSheetView

#pragma mark - init

/* 
 Creates a UIScrollview & UITableView. 
 - The UIScrollView will help in horizontal scroll 
 - The UITableView will be responsible for creating rows & vertical scroll. 
 
 - The number of items in columnData will decide the number of columns of XLS Sheet.
 - Each columnData node to have the property key (along with all other metadata info) of that 
   column which will be looked into the xlsData objects.
 - Each xlsData object should be array of NSDictionary or any NSObject class having the all required 
   properties for each column.
 */
-(id) initWithFrame:(CGRect)frame
               Data:(NSArray*) xlsData
            Columns:(NSArray*) columnData
      TotalProperty:(NSArray*) totalProperties
       HeaderHeight:(CGFloat) headerHeight
          RowHeight:(CGFloat) rowHeight
    WithDefaultSort:(NSString*) sortproperty
          SortOrder:(UIExcelSortOrder)order
{
    __LOG__
    
    self = [super initWithFrame:frame];
    
    if (self)
    {
        self.rowHeight = rowHeight;
        self.headerHeight = headerHeight;
        
        self.totalRowData = totalProperties;
        self.columnData = columnData;
        self.xlsData = [xlsData mutableCopy];
        
        CGRect scrollFrame = CGRectZero;
        scrollFrame.size = frame.size;
        self.tableScrollView = [[UIScrollView alloc] initWithFrame:scrollFrame];
        
        [self.tableScrollView setBackgroundColor:[UIColor clearColor]];
        
        [self setBackgroundColor:[UIColor colorWithPatternImage:[UIImage imageNamed:@"backgroundGray.jpg"]]];
        
        [self addSubview:self.tableScrollView];
        
        self.tableView = [[UITableView alloc] initWithFrame:scrollFrame style:UITableViewStylePlain];
        self.tableView.dataSource = self;
        self.tableView.delegate = self;
        [self.tableView setBackgroundColor:[UIColor clearColor]];
        [self.tableView setSeparatorColor:[UIColor lightGrayColor]];
        //        [self.tableView setScrollEnabled:TRUE];
        
        [self.tableScrollView addSubview:self.tableView];
        
        CGFloat totalWidth = [self totalWidth];
        CGFloat totalHeight = [self totalHeight];
        
        CGFloat viewHeight = self.tableScrollView.frame.size.width;
        
        
        if (totalHeight < viewHeight)
            viewHeight = totalHeight;
        
        [self.tableView setFrame:CGRectMake(0, 0, totalWidth, viewHeight)];
        
        CGSize contentSize = CGSizeMake(totalWidth,viewHeight);
        
        [self.tableScrollView setContentSize:contentSize];
        
        if (sortproperty)
        {
            self.sortProperty = sortproperty;
        }
        else {
            for (ExcelSheetColumnData* column in self.columnData)
            {
                if ([column.propertyKey isEqualToString:[NSString stringWithFormat:@"%@",kExcelSheetIndexColumn]])
                    continue;
                else
                {
                    self.sortProperty = column.propertyKey;
                    break;
                }
            }
        }
        
        self.sortOrder = order;
        
        [self sortData];
    }
    
    return self;
}



-(id) initWithFrame:(CGRect) frame
               Data:(NSArray*) xlsData
            Columns:(NSArray*) columnData
    WithDefaultSort:(NSString*) sortproperty
          SortOrder:(UIExcelSortOrder) order
{
    __LOG__
    
    return [self initWithFrame:frame
                          Data:xlsData
                       Columns:columnData
                  HeaderHeight:HEADER_HEIGHT
                     RowHeight:ROW_HEIGHT
               WithDefaultSort:sortproperty
                     SortOrder:order];
}

-(id) initWithFrame:(CGRect)frame
               Data:(NSArray*) xlsData
            Columns:(NSArray*) columnData
       HeaderHeight:(CGFloat) headerHeight
          RowHeight:(CGFloat) rowHeight
    WithDefaultSort:(NSString*) sortproperty
          SortOrder:(UIExcelSortOrder)order
{
    __LOG__
    
    return [self initWithFrame:frame
                          Data:xlsData
                       Columns:columnData
                 TotalProperty:nil
                  HeaderHeight:headerHeight
                     RowHeight:rowHeight
               WithDefaultSort:sortproperty
                     SortOrder:order];
}

-(void) setTotalRowData:(NSArray *)totalRowData
{
    _totalRowData = totalRowData;
    
    // if total properties are set, add one blank coloumn in the first index and label the
    // cell with "Total" after all the data is shown
    if ([self.xlsData count])
        [self.tableView reloadData];
}

#pragma mark -

-(void) ReplaceObjectAtIndex:(int) index WithObject:(id) object
{
    if (index < [self.xlsData count])
    {
        [self.xlsData replaceObjectAtIndex:index withObject:object];
        
        [self.tableView reloadRowsAtIndexPaths:@[[NSIndexPath indexPathForRow:index inSection:0]] withRowAnimation:UITableViewRowAnimationNone];
    }
}

-(void) RefreshSheetWithData:(NSArray*) xlsData Columns:(NSArray*) columnData
{
    self.xlsData = [xlsData mutableCopy];
    self.columnData = columnData;
    
    [self sortData];
}

-(void) SetSortByProperty:(NSString *)sortedByProperty Order:(UIExcelSortOrder)sortOrder
{
    self.sortProperty = sortedByProperty;
    self.sortOrder = sortOrder;
    [self sortData];
}

-(void) sortData
{
    if (![self.sortProperty length])
        return;
    
    ExcelSheetColumnData* columnData = nil;
    
    for (ExcelSheetColumnData* data in self.columnData)
    {
        if ([data.propertyKey isEqualToString:self.sortProperty])
        {
            columnData = data;
            break;
        }
    }
    
    [self.xlsData sortUsingComparator:^NSComparisonResult(id obj1, id obj2) {
        NSComparisonResult retValue = NSOrderedSame;
        
        if (columnData.type == ColumnDataType_RegularString)
        {
            NSString* str1 = [obj1 valueForKey:self.sortProperty];
            NSString* str2 = [obj2 valueForKey:self.sortProperty];
            
            if (self.sortOrder == UIExcelSortOrderAscending) {
                retValue = [str1 compare:str2];
            }
            else {
                retValue = [str2 compare:str1];
            }
        }
        else if(columnData.type == ColumnDataType_Date)
        {
            NSDate* date1 = [obj1 valueForKey:self.sortProperty];
            NSDate* date2 = [obj2 valueForKey:self.sortProperty];
            
            if (self.sortOrder == UIExcelSortOrderAscending) {
                retValue = [date1 compare:date2];
            }
            else {
                retValue = [date2 compare:date1];
            }
        }
        else if(columnData.type == ColumnDataType_Number)
        {
            NSNumber* num1 = [obj1 valueForKey:self.sortProperty];
            NSNumber* num2 = [obj2 valueForKey:self.sortProperty];
            
            if (self.sortOrder == UIExcelSortOrderAscending) {
                retValue = [num1 compare:num2];
            }
            else {
                retValue = [num2 compare:num1];
            }
            
        }
        
        return retValue;
    }];
    
    [self.tableView reloadData];
}

-(void) UpdateRows:(NSArray*) indexPaths
{
    [self.tableView reloadRowsAtIndexPaths:indexPaths withRowAnimation:UITableViewRowAnimationNone];
}

-(CGFloat) totalHeight
{
    CGFloat height = 0;
    
    int rowCount = [self.xlsData count];
    
    if([self.totalRowData count])
        rowCount++;
    
    height = (rowCount * self.rowHeight) + self.headerHeight;
    
    return height;
}

-(CGFloat) totalWidth
{
    CGFloat width = 0;
    
    for (ExcelSheetColumnData* colData in self.columnData) {
        width += colData.width;
    }
    
    return width;
}

/* calculates the total amount to display in the "Total" row at the end.
 - The calculation will be based on the type of total mentioned in the total column data.
 */
-(NSDictionary*) CalculateTotalRowData
{
    NSMutableDictionary* totalData = [NSMutableDictionary new];
    
    for (int iter = 0; iter < [self.columnData count]; iter++)
    {
        ExcelSheetColumnData* column = [self.columnData objectAtIndex:iter];

        NSMutableDictionary* totalDictionary = [NSMutableDictionary new];
        [totalDictionary setObject:@(iter) forKey:@"columnIndex"];
        
        [totalData setObject:totalDictionary forKey:column.propertyKey];
    }
    
    for (id object in self.xlsData)
    {
        for (ExcelSheetTotalRowData* totalInfo in self.totalRowData)
        {
            NSNumber* number = [object valueForKey:totalInfo.property];
            
            if (totalInfo.totalType == ExcelSheetTotalData_Sum)
            {
                if (![number isKindOfClass:[NSNumber class]])
                    continue;
                
                float numberIntValue = [number floatValue];
                
                NSMutableDictionary* totalPropertyData = [totalData objectForKey:totalInfo.property];
                
                float totalAmount = [[totalPropertyData objectForKey:@"totalAmount"] floatValue];
                totalAmount += numberIntValue;
                [totalPropertyData setObject:@(totalAmount) forKey:@"totalAmount"];
            }
            else if(totalInfo.totalType == ExcelSheetTotalData_Count)
            {
                NSMutableDictionary* totalPropertyData = [totalData objectForKey:totalInfo.property];
                
                int totalCount = [[totalPropertyData objectForKey:@"totalCount"] integerValue];
                totalCount++;
                [totalPropertyData setObject:@(totalCount) forKey:@"totalCount"];
            }
        }
    }
    
    return totalData;
}

-(UITableViewCell*) ItemAtIndexPath:(NSIndexPath*) indexPath
{
    if (indexPath)
        return [self.tableView cellForRowAtIndexPath:indexPath];
    
    return  nil;
}

-(CGRect) RectForCellForIndexPath:(NSIndexPath*) indexPath
{
    UITableViewCell* cell = [self.tableView cellForRowAtIndexPath:indexPath];
    return cell.frame;
}

-(void) SetHeaderRowHeight:(CGFloat) headerHeight RowHeight:(CGFloat) rowHeight
{
    self.headerHeight = headerHeight;
    self.rowHeight = rowHeight;
    
    // reload the table if already loaded.
    if ([self.tableView numberOfRowsInSection:0]) {
        [self.tableView reloadData];
    }
}

#pragma mark - table view API

-(int) numberOfSectionsInTableView:(UITableView *)tableView
{
    __LOG__
    return 1;
}

-(CGFloat) tableView:(UITableView *)tableView heightForRowAtIndexPath:(NSIndexPath *)indexPath
{
    __LOG__
    return self.rowHeight;
}

- (NSInteger)tableView:(UITableView *)tableView numberOfRowsInSection:(NSInteger)section
{
    __LOG__
    
    NSInteger numRows = 0;
    
    // add one to display "Total" if totalRowData is valid
    if ([self.totalRowData count])
        numRows = [self.xlsData count] + 1;
    else
        numRows = [self.xlsData count];
    
    return numRows;
}

-(CGFloat) tableView:(UITableView *)tableView heightForHeaderInSection:(NSInteger)section
{
    __LOG__
    return self.headerHeight;
}

/* 
 Configures the header row. 
 - Each header cell will have to a button so that clicking can be handled to change the sort property.
 - It will also have an image to indicate sorting order.
 */
-(UIView*) tableView:(UITableView *)tableView viewForHeaderInSection:(NSInteger)section
{
    __LOG__
    UIView* headerView = [[UIView alloc] initWithFrame:CGRectMake(0, 0, tableView.frame.size.width, self.headerHeight)];
    // r,g,b = 245,211, 49
    [headerView setBackgroundColor:[UIColor colorWithRed:(245.0f/255.0f)
                                                   green:(211.0f/255.0f)
                                                    blue:(49.0f/255.0f) alpha:1.0]];
    
    UIButton* button;
    UILabel* label;
    UIImageView* sortImage;
    
    int tag = 0;
    int iter = 0;
    int lineStartX = 0;
    
    CGRect headerRect = CGRectMake(0, 0, 0, self.headerHeight);
    CGRect sortImageRect = CGRectMake(0, 0, 20, self.rowHeight);
    
    // a header has Button, label & sort image per column.
    // it also has a separator vertical line per column.
    
    for (ExcelSheetColumnData* colData in self.columnData)
    {
        // set button first.
        headerRect.size.width = (colData.width - (RIGHT_PADDING+LEFT_PADDING));
        
        tag = (COLUMNSTART_TAG + iter);
        
        button = (UIButton*)[headerView viewWithTag:tag];
        
        if (!button)
        {
            button = [UIButton buttonWithType:UIButtonTypeCustom];
            [button setTag:tag];
            [headerView addSubview:button];
            [button addTarget:self action:@selector(headerButtonHandler:) forControlEvents:UIControlEventTouchUpInside];
        }
        
        [button setFrame:headerRect];
        
        // set label
        label = (UILabel*)[button viewWithTag:1];
        
        if (!label)
        {
            label = [UILabel new];
            [label setTag:1];
            [button addSubview:label];
            [label setAdjustsFontSizeToFitWidth:TRUE];
            [label setMinimumScaleFactor:0.5];
        }
        
        [label setBackgroundColor:[UIColor clearColor]];
        [label setTextAlignment:NSTextAlignmentCenter];
        [label setFont:[UIFont boldSystemFontOfSize:DEFAULT_HEADER_FONTSIZE]];
        
        CGRect labelRect = headerRect;
        labelRect.size.width = headerRect.size.width - (LEFT_PADDING + sortImageRect.size.width + 2);
        labelRect.origin = CGPointMake(LEFT_PADDING, 0);
        
        [label setFrame:labelRect];
        
        // set sort image if needed
        sortImage = (UIImageView*)[button viewWithTag:2];
        
        if (!sortImage) {
            sortImage = [UIImageView new];
            [sortImage setTag:2];
            [sortImage setBackgroundColor:[UIColor clearColor]];
            [sortImage setContentMode:UIViewContentModeScaleAspectFit];
            [button addSubview:sortImage];
        }
        
        sortImageRect.origin.x = labelRect.origin.x + labelRect.size.width + 2;
        [sortImage setFrame:sortImageRect];
        
        if ([self.sortProperty length])
        {
            if ([self.sortProperty isEqualToString:colData.propertyKey])
            {
                if (self.sortOrder == UIExcelSortOrderAscending)
                {
                    [sortImage setImage:[UIImage imageNamed:@"sortAscending.png"]];
                }
                else {
                    [sortImage setImage:[UIImage imageNamed:@"sortDescending.png"]];
                }
            }
            else {
                [sortImage setImage:nil];
            }
        }
        else {
            [sortImage setImage:nil];
        }
        
        // put vertical line
        lineStartX = headerRect.origin.x + colData.width - LEFT_PADDING;
        CALayer* line = [[CALayer alloc] init];
        [line setBackgroundColor:[[UIColor blackColor] CGColor]];
        [line setFrame:CGRectMake(lineStartX, 0, 1, self.headerHeight)];
        [headerView.layer insertSublayer:line atIndex:0];
        
        headerRect.origin.x += colData.width;
        
        [label setText:colData.columnTitle ? colData.columnTitle : @""];
        
        iter++;
    }
    
    return headerView;
}

/*
 Configure each row.
 - This function will take care of creating cells for each row in the XLS sheet.
 - Each cell will have a text label which can be configured to display a specific font and coloring.
 */
- (UITableViewCell *)tableView:(UITableView *)tableView
         cellForRowAtIndexPath:(NSIndexPath *)indexPath
{
    UITableViewCell* cell = [tableView dequeueReusableCellWithIdentifier:@"tableViewCell"];
    
    if (!cell)
    {
        cell = [[UITableViewCell alloc] initWithStyle:UITableViewCellStyleDefault reuseIdentifier:@"tableViewCell"];
    }
    
    int iter = 0;
    int lineStartX = 0;
    CGRect cellViewRect = CGRectMake(0, 0, 0, self.rowHeight);
    CGRect labelRect = cellViewRect;
    labelRect.origin.x = LEFT_PADDING;
    
    if ([indexPath row] < [self.xlsData count])
    {
        id object = [self.xlsData objectAtIndex:[indexPath row]];
        int tag = 0;
        
        UIView* cellView;
        UILabel* label;
        
        for (ExcelSheetColumnData* colData in self.columnData)
        {
            tag = (COLUMNSTART_TAG + iter);
            
            cellViewRect.size.width = colData.width;
            
            if (!cellViewRect.size.width)
                cellViewRect.size.width = 100;
            
            cellView = (UIView*) [cell viewWithTag:tag];
            
            if (!cellView) {
                cellView = [[UIView alloc] init];
                [cellView setTag:tag];
                [cell addSubview:cellView];
            }
            
            [cellView setBackgroundColor:[UIColor whiteColor]];
            [cellView setFrame:cellViewRect];
            
            label = (UILabel*)[cellView viewWithTag:CELL_LABEL_TAG];
            
            if (!label)
            {
                label = [UILabel new];
                [label setTag:CELL_LABEL_TAG];
                [cellView addSubview:label];
                
                lineStartX = cellViewRect.size.width - LEFT_PADDING;
                CALayer* line = [[CALayer alloc] init];
                [line setBackgroundColor:[[UIColor blackColor] CGColor]];
                [line setFrame:CGRectMake(lineStartX, 0, 1, self.rowHeight)];
                [cellView.layer insertSublayer:line atIndex:0];
            }
            
            [label setBackgroundColor: colData.backColor ? colData.backColor : [UIColor whiteColor]];
            [label setLineBreakMode:NSLineBreakByTruncatingMiddle];
            [label setTextAlignment:NSTextAlignmentCenter];
            
            if (colData.customFont) {
                [label setFont:colData.customFont];
            }
            else {
                if (colData.formatBold && !colData.formatItalic) {
                    [label setFont:[UIFont boldSystemFontOfSize:DEFAULT_CELL_FONTSIZE]];
                }
                else if(!colData.formatBold && colData.formatItalic) {
                    [label setFont:[UIFont italicSystemFontOfSize:DEFAULT_CELL_FONTSIZE]];
                }
                else if(colData.formatItalic && colData.formatBold)
                {
                    [label setFont:[UIFont boldSystemFontOfSize:DEFAULT_CELL_FONTSIZE]];
                }
                else {
                    [label setFont:[UIFont systemFontOfSize:DEFAULT_CELL_FONTSIZE]];
                }
            }
            
            [label setTextColor:colData.fontColor ? colData.fontColor : [UIColor blackColor]];
            
            
            labelRect.size.width = (colData.width - (LEFT_PADDING + RIGHT_PADDING));
            [label setFrame:labelRect];
            
            if (colData.type == ColumnDataType_Date)
            {
                if (colData.dateDisplayType == DateTimeDisplay_TimeSinceNow)
                {
                    NSDate* date = [object valueForKey:colData.propertyKey];
                    NSDateFormatter* dateFormatter = [[NSDateFormatter alloc] init];
                    [dateFormatter setDateStyle:NSDateFormatterMediumStyle];
                    [dateFormatter setTimeStyle:NSDateFormatterMediumStyle];
                    
                    NSString* displayString = [dateFormatter stringFromDate:date];
                    [label setText:displayString];
                }
                else
                {
                    NSDateFormatter* dateFormatter = [[NSDateFormatter alloc] init];
                    [dateFormatter setDateStyle:NSDateFormatterMediumStyle];
                    [dateFormatter setTimeStyle:NSDateFormatterShortStyle];
                    
                    NSDate* date = [object valueForKey:colData.propertyKey];
                    NSString* dateString = [dateFormatter stringFromDate:date];
                    [label setText:dateString];
                }
            }
            else {
                if ([colData.propertyKey isEqualToString:[NSString stringWithFormat:@"%@",kExcelSheetIndexColumn]]) {
                    int rowNum = ([indexPath row] + 1);
                    [label setText:[@(rowNum) description]];
                }
                else {
                    [label setText:[[object valueForKey:colData.propertyKey] description]];
                }
            }
            
            cellViewRect.origin.x += colData.width;
            
            iter++;
        }
    }
    else
    {
        // this is a total row
        NSDictionary* totalData = [self CalculateTotalRowData];
        
        int tag = 0;
        
        UIView* cellView;
        UILabel* label;
        
        for (ExcelSheetColumnData* colData in self.columnData)
        {
            tag = (COLUMNSTART_TAG + iter);
            
            cellViewRect.size.width = colData.width;
            
            cellView = (UIView*) [cell viewWithTag:tag];
            
            if (!cellView) {
                cellView = [[UIView alloc] init];
                [cellView setTag:tag];
                [cell addSubview:cellView];
            }
            
            [cellView setBackgroundColor:[UIColor lightGrayColor]];
            [cellView setFrame:cellViewRect];
            
            label = (UILabel*)[cellView viewWithTag:CELL_LABEL_TAG];
            
            if (!label)
            {
                label = [UILabel new];
                [label setTag:CELL_LABEL_TAG];
                [cellView addSubview:label];
                
                lineStartX = cellViewRect.size.width - LEFT_PADDING;
                CALayer* line = [[CALayer alloc] init];
                [line setBackgroundColor:[[UIColor blackColor] CGColor]];
                [line setFrame:CGRectMake(lineStartX, 0, 1, self.rowHeight)];
                [cellView.layer insertSublayer:line atIndex:0];
            }
            
            [label setBackgroundColor: [UIColor clearColor]];
            [label setTextAlignment:NSTextAlignmentCenter];
            [label setFont:[UIFont boldSystemFontOfSize:DEFAULT_CELL_FONTSIZE]];
            [label setTextColor:[UIColor blackColor]];
            
            labelRect.size.width = (colData.width - (LEFT_PADDING + RIGHT_PADDING));
            [label setFrame:labelRect];
            
            if ([colData.propertyKey isEqualToString:[NSString stringWithFormat:@"%@",kExcelSheetIndexColumn]]) {
                [label setText:@"Total"];
            }
            else {
                NSDictionary* totalPropertyData = [totalData objectForKey:colData.propertyKey];
                NSNumber* num = nil;
                
                if ([totalPropertyData objectForKey:@"totalAmount"]) {
                    num = [totalPropertyData objectForKey:@"totalAmount"];
                }
                else if([totalPropertyData objectForKey:@"totalCount"]) {
                    num = [totalPropertyData objectForKey:@"totalCount"];
                }
                
                if (num)
                    [label setText:[num description]];
                else
                    [label setText:@""];
            }
            
            cellViewRect.origin.x += colData.width;
            
            iter++;
        }
    }
    
    return cell;
}

-(void) tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath
{
    if (!self.stickySelection)
        [tableView deselectRowAtIndexPath:indexPath animated:YES];
    
    if ([indexPath row] < [self.xlsData count] && [self.delegate respondsToSelector:@selector(UIExcelSheetView:DidSelectItem:AtIndexPath:)])
    {
        id itemData = [self.xlsData objectAtIndex:[indexPath row]];
        [self.delegate UIExcelSheetView:self DidSelectItem:itemData AtIndexPath:indexPath];
    }
}

#pragma mark - Button handlers

-(void) headerButtonHandler:(UIButton*) button
{
    int columnIndex = (button.tag - COLUMNSTART_TAG);
    
    if (columnIndex >= 0 && columnIndex < [self.columnData count])
    {
        ExcelSheetColumnData* columnData = [self.columnData objectAtIndex:columnIndex];
        
        UIExcelSortOrder sortOrder = UIExcelSortOrderAscending;
        
        if ([self.sortProperty isEqualToString:columnData.propertyKey])
        {
            if (self.sortOrder == UIExcelSortOrderAscending) {
                sortOrder = UIExcelSortOrderDescending;
            }
            else {
                sortOrder = UIExcelSortOrderAscending;
            }
        }
        
        [self SetSortByProperty:columnData.propertyKey Order:sortOrder];
    }
}

-(void) columnDataButtonHandler:(UIButton*) button
{
    UITableViewCell* cell = (UITableViewCell*)[[button superview] superview];
    NSIndexPath* indexPath = [self.tableView indexPathForCell:cell];
    
    if ([indexPath row] < [self.xlsData count])
    {
        id itemData = [self.xlsData objectAtIndex:[indexPath row]];
        
        [self.delegate UIExcelSheetView:self
                                  Event:Event_ButtonClicked
                            ForItemData:itemData
                               UserInfo:button];
    }
    
}

@end
