//
//  PSCell.h
//  PSXLSReader
//
//  Created by Pan on 2018/5/24.
//

#import <Foundation/Foundation.h>

typedef NS_ENUM(NSUInteger, PSCellContentType) {
    PSCellContentTypeBlank=0,
    PSCellContentTypeString,
    PSCellContentTypeInteger,
    PSCellContentTypeFloat,
    PSCellContentTypeBool,
    PSCellContentTypeError,
    PSCellContentTypeUnknown
};


@interface PSCell : NSObject

@property (nonatomic, assign, readonly) PSCellContentType type;
@property (nonatomic, assign, readonly) NSInteger row;
@property (nonatomic, assign, readonly) NSInteger col;
@property (nonatomic, assign, readonly) NSString *columnName;            // "A" ... "Z", "AA"..."ZZZ"
@property (nonatomic, strong, readonly) NSString *str;        // typeof depends on contentsType
@property (nonatomic, strong, readonly) NSNumber *val;        // typeof depends on contentsType

+ (PSCell *)blankCell;

@end
