//
//  PSXLSReader.h
//  PSXLSReader
//
//  Created by Pan on 2018/5/24.
//

#import <Foundation/Foundation.h>
#import "PSCell.h"

NS_ASSUME_NONNULL_BEGIN

@interface PSXLSReader : NSObject

// Summary information of xls file
@property (nonatomic, readonly) NSString *appName;
@property (nonatomic, readonly) NSString *author;
@property (nonatomic, readonly) NSString *category;
@property (nonatomic, readonly) NSString *comment;
@property (nonatomic, readonly) NSString *company;
@property (nonatomic, readonly) NSString *keywords;
@property (nonatomic, readonly) NSString *lastAuthor;
@property (nonatomic, readonly) NSString *manager;
@property (nonatomic, readonly) NSString *subject;
@property (nonatomic, readonly) NSString *title;

+ (instancetype)readerWithPath:(NSString *)filePath;

- (NSString *)libaryVersion;

// Sheet Information
- (NSInteger)numberOfSheets;
- (NSString *)sheetNameAtIndex:(NSInteger)index;
- (NSInteger)rowsForSheetAtIndex:(NSInteger)idx;
- (BOOL)isSheetVisibleAtIndex:(NSUInteger)index;
- (NSInteger)numberOfRowsInSheet:(NSInteger)sheetIndex;
- (NSInteger)numberOfColsInSheet:(NSInteger)sheetIndex;

// Random Access
- (PSCell *)cellInWorkSheetIndex:(NSInteger)sheetNum row:(NSInteger)row column:(NSInteger)column;        // uses 1 based indexing!
- (PSCell *)cellInWorkSheetIndex:(NSInteger)sheetNum row:(NSInteger)row columnString:(NSString *)column;        // "A"...."Z" "AA"..."ZZ"

// Iterate through all cells
- (void)startIteratorSheetAtIndex:(NSInteger)sheetIndex;
- (nullable PSCell *)nextCell;

@end

NS_ASSUME_NONNULL_END
