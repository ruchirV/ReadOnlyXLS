//
//  UIExcelSheetView.h
//  MyExcel
//
//  Created by Ruchir on 12/05/14.
//  Copyright (c) 2014 yTrre Solutions Pvt Ltd.. All rights reserved.
//

#import <UIKit/UIKit.h>

/* 
  - The sorting type. default will be Ascending. The sorting can be changed by clickin on the
    header cell.
 */
typedef enum
{
    UIExcelSortOrderAscending = 0,
    UIExcelSortOrderDescending,
} UIExcelSortOrder;

/* We can keep native UI controls in a cell. Not fully implemented yet. */
typedef enum
{
    Event_ButtonClicked,
    Event_SwitchValueChanged,
} ExcelSheetEvent;

@class UIExcelSheetView;

@protocol UIExcelSheetViewDelegate <NSObject>

@optional

/* called when a row is selected */
-(void)UIExcelSheetView:(UIExcelSheetView*) xlsSheet
          DidSelectItem:(id) itemData
            AtIndexPath:(NSIndexPath*) indexPath;

/* called when a UI component like UIButton or UISwitch is clicked */
-(void)UIExcelSheetView:(UIExcelSheetView*) xlsSheet
                  Event:(ExcelSheetEvent) event
            ForItemData:(id) itemData
               UserInfo:(id) userInfo;
@end

@interface UIExcelSheetView : UIView <UITableViewDataSource,
UITableViewDelegate>

@property (nonatomic, weak) id<UIExcelSheetViewDelegate> delegate;
@property (nonatomic, assign) BOOL stickySelection;

/* Call init and pass a frame to it. Andy frame can be given. Currently the UI adaptability is
   not supported. 
 
 @frame: the frame of the XLS Sheet
 @xlsData: Either array of dictionary or class objects.
 @columnData: Array of ExcelSheetColumnData objects.
 @totalProperties: The array of the property key of the columsn for which total is requried.
 @rowHeight: any specific height for a row.
 @sortProperty: the property key of the column for which sorting should be done and displayed. 
 @order: Ascending or Descending.
 */
-(id) initWithFrame:(CGRect)frame
               Data:(NSArray*) xlsData
            Columns:(NSArray*) columnData
      TotalProperty:(NSArray*) totalProperties
       HeaderHeight:(CGFloat) headerHeight
          RowHeight:(CGFloat) rowHeight
    WithDefaultSort:(NSString*) sortproperty
          SortOrder:(UIExcelSortOrder)order;

/* Call this when the data being displayed is refreshed due to local fetch or otherwise. */
-(void) RefreshSheetWithData:(NSArray*) xlsData Columns:(NSArray*) columnData;

/* Update individual rows */
-(void) UpdateRows:(NSArray*) indexPaths;

/* Call this if some specific column sorting is needed */
-(void) SetSortByProperty:(NSString*) sortedByProperty
                    Order:(UIExcelSortOrder) sortOrder;

/* useful if rect is needed to display any UIPopover */
-(CGRect) RectForCellForIndexPath:(NSIndexPath*) indexPath;

/* useful if we have to change the data shown at an index */
-(void) ReplaceObjectAtIndex:(int) index WithObject:(id) object;

/* if the header & regular row heights need to be changed. */
-(void) SetHeaderRowHeight:(CGFloat) headerHeight RowHeight:(CGFloat) rowHeight;

/* returns the cell pointer for the mentioned index path */
-(UITableViewCell*) ItemAtIndexPath:(NSIndexPath*) indexPath;

@end
