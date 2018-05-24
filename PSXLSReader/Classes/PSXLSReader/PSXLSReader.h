//
//  PSXLSReader.h
//  PSXLSReader
//
//  Created by Pan on 2018/5/24.
//

#import <Foundation/Foundation.h>
#import "PSCell.h"

enum {DHWorkSheetNotFound = UINT32_MAX};

@interface PSXLSReader : NSObject

+ (PSXLSReader *)readerWithPath:(NSString *)filePath;

- (NSString *)libaryVersion;

// Sheet Information
- (uint32_t)numberOfSheets;
- (NSString *)sheetNameAtIndex:(uint32_t)index;
- (uint16_t)rowsForSheetAtIndex:(uint32_t)idx;
- (BOOL)isSheetVisibleAtIndex:(NSUInteger)index;
- (uint16_t)numberOfRowsInSheet:(uint32_t)sheetIndex;
- (uint16_t)numberOfColsInSheet:(uint32_t)sheetIndex;

// Random Access
- (PSCell *)cellInWorkSheetIndex:(uint32_t)sheetNum row:(uint16_t)row col:(uint16_t)col;        // uses 1 based indexing!
- (PSCell *)cellInWorkSheetIndex:(uint32_t)sheetNum row:(uint16_t)row colStr:(char *)col;        // "A"...."Z" "AA"..."ZZ"

// Iterate through all cells
- (void)startIterator:(uint32_t)sheetNum;
- (PSCell *)nextCell;

// Summary Information
- (NSString *)appName;
- (NSString *)author;
- (NSString *)category;
- (NSString *)comment;
- (NSString *)company;
- (NSString *)keywords;
- (NSString *)lastAuthor;
- (NSString *)manager;
- (NSString *)subject;
- (NSString *)title;

@end
