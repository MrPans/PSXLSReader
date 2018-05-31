//
//  PSXLSCell.m
//  PSXLSReader
//
//  Created by Pan on 2018/5/24.
//

#import "PSXLSCell.h"

@interface PSXLSCell()
@property (nonatomic, assign, readwrite) PSCellContentType type;
@property (nonatomic, assign, readwrite) NSInteger row;
@property (nonatomic, assign, readwrite) NSInteger column;
@property (nonatomic, assign, readwrite) NSString *columnName;            // "A" ... "Z", "AA"..."ZZZ"
@property (nonatomic, strong, readwrite) NSString *text;        // typeof depends on contentsType
@property (nonatomic, strong, readwrite) NSNumber *value;        // typeof depends on contentsType
@end

@implementation PSXLSCell

+ (PSXLSCell *)blankCell
{
    return [[PSXLSCell alloc] init];
}

- (NSString *)description {
    
    NSString *pointerInformation = [super description];
    NSDictionary *nameMapper = @{
                                 @(PSCellContentTypeBlank): @"Blank",
                                 @(PSCellContentTypeString): @"String",
                                 @(PSCellContentTypeInteger): @"Integer",
                                 @(PSCellContentTypeFloat): @"Float",
                                 @(PSCellContentTypeBool): @"Bool",
                                 @(PSCellContentTypeError): @"Error",
                                 };
    NSDictionary *dic = @{
                          @"type": nameMapper[@(self.type)],
                          @"row": @(self.row),
                          @"column": @(self.column),
                          @"columnName": self.columnName,
                          @"text": self.text,
                          @"value": self.value,
                          };
    NSString *desc = [NSString stringWithFormat:@"\n%@\n%@\n", pointerInformation, dic];
    return desc;
}
@end
