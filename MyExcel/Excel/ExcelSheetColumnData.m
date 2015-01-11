//
//  ExcelSheetColumnData.m
//  yTrre Merchant
//
//  Created by Ruchir on 02/05/14.
//

#import "ExcelSheetColumnData.h"

const NSString* kExcelSheetIndexColumn = @"ExcelSheetIndexColumn";

@implementation ExcelSheetColumnData

/* default index. Just call this function and add the column in the column array at first index */
+(ExcelSheetColumnData*) IndexColumn
{
    ExcelSheetColumnData* colData = [ExcelSheetColumnData new];
    colData.propertyKey = [NSString stringWithString:[NSString stringWithFormat:@"%@",kExcelSheetIndexColumn]];
    colData.columnTitle = @"Sr No";
    colData.width = 80;
    colData.formatBold = TRUE;
    return colData;
}

@end

/* no implementation needed */
@implementation ExcelSheetTotalRowData

@end
