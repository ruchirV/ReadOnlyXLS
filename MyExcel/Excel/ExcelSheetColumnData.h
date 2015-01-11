//
//  ExcelSheetColumnData.h
//  MyExcel
//
//  Created by Ruchir on 02/05/14.
//

/*
 Basic component of the xls sheet.
 */

#import <Foundation/Foundation.h>

/* useful, specially when sorting is done. */
typedef enum
{
    ColumnDataType_RegularString = 0, /* used for regular string */
    ColumnDataType_Number, /* used to display integers, floats, etc */
    ColumnDataType_Date, /* used to diplsay date type of objects. */
} ExcelSheetColumnDataType;

/* date display types */
typedef enum
{
    DateDisplay_ShortType,
    DateDisplay_MediumType,
    DateDisplay_LongType,
    DateTimeDisplay_ShortType,
    DateTimeDisplay_MediumType,
    DateTimeDisplay_LongType,
    DateTimeDisplay_TimeSinceNow,
}DateDisplayType;

const NSString* kExcelSheetIndexColumn;

@interface ExcelSheetColumnData : NSObject

@property (nonatomic, copy)     NSString*                   propertyKey;
@property (nonatomic, copy)     NSString*                   columnTitle;
@property (nonatomic, assign)   ExcelSheetColumnDataType    type;
@property (nonatomic, assign)   DateDisplayType             dateDisplayType;
@property (nonatomic, assign)   int                         width;
@property (nonatomic, assign)   BOOL                        formatBold;
@property (nonatomic, assign)   BOOL                        formatItalic;
@property (nonatomic, strong)   UIColor*                    fontColor;
@property (nonatomic, strong)   UIColor*                    backColor;
@property (nonatomic, strong)   UIFont*                     customFont;

+(ExcelSheetColumnData*) IndexColumn;

@end

typedef enum
{
    ExcelSheetTotalData_Sum = 0,
    ExcelSheetTotalData_Count,
} ExcelSheetTotalDataType;

@interface ExcelSheetTotalRowData : NSObject

@property (nonatomic, copy) NSString* property;
@property (nonatomic, assign) int totalType;

@end
