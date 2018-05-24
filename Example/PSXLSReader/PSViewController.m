//
//  PSViewController.m
//  PSXLSReader
//
//  Created by DeveloperPans on 05/24/2018.
//  Copyright (c) 2018 DeveloperPans. All rights reserved.
//

#import "PSViewController.h"
@import PSXLSReader;

@interface PSViewController ()

@end

@implementation PSViewController

- (void)viewDidLoad
{
    [super viewDidLoad];
    [self test];
}

- (void)test {
    [super viewDidLoad];
    
    NSString *path = [[[NSBundle mainBundle] resourcePath] stringByAppendingPathComponent:@"test.xls"];
    
    PSXLSReader *reader = [PSXLSReader readerWithPath:path];
    
    NSString *text = @"";
    
    text = [text stringByAppendingFormat:@"AppName: %@\n", reader.appName];
    text = [text stringByAppendingFormat:@"Author: %@\n", reader.author];
    text = [text stringByAppendingFormat:@"Category: %@\n", reader.category];
    text = [text stringByAppendingFormat:@"Comment: %@\n", reader.comment];
    text = [text stringByAppendingFormat:@"Company: %@\n", reader.company];
    text = [text stringByAppendingFormat:@"Keywords: %@\n", reader.keywords];
    text = [text stringByAppendingFormat:@"LastAuthor: %@\n", reader.lastAuthor];
    text = [text stringByAppendingFormat:@"Manager: %@\n", reader.manager];
    text = [text stringByAppendingFormat:@"Subject: %@\n", reader.subject];
    text = [text stringByAppendingFormat:@"Title: %@\n", reader.title];
    
    text = [text stringByAppendingFormat:@"\n\nNumber of Sheets: %u\n", reader.numberOfSheets];
    NSLog(@"%@", text);
    [reader startIterator:0];
    
    while(YES) {
        PSCell *cell = [reader nextCell];
        if(cell.type == PSCellContentTypeBlank) break;
        NSLog(@"%@", cell);
    }
}

@end
