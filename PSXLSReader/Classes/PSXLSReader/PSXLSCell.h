//
//  PSXLSCell.h
//  PSXLSReader
//
//  Created by Pan on 2018/5/24.
//

#import <Foundation/Foundation.h>

NS_ASSUME_NONNULL_BEGIN

typedef NS_ENUM(NSUInteger, PSCellContentType) {
    PSCellContentTypeBlank = 0,
    PSCellContentTypeString,
    PSCellContentTypeInteger,
    PSCellContentTypeFloat,
    PSCellContentTypeBool,
    PSCellContentTypeError,
    PSCellContentTypeUnknown
};


@interface PSXLSCell : NSObject

@property (nonatomic, assign, readonly) PSCellContentType type;
@property (nonatomic, assign, readonly) NSInteger row;
@property (nonatomic, assign, readonly) NSInteger column;
@property (nonatomic, assign, readonly) NSString *columnName;            // "A" ... "Z", "AA"..."ZZZ"
@property (nonatomic, strong, readonly) NSString *text;        // typeof depends on contentsType
@property (nonatomic, strong, readonly) NSNumber *value;        // typeof depends on contentsType

+ (instancetype)blankCell;

@end

NS_ASSUME_NONNULL_END
