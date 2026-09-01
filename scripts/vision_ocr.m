#import <AppKit/AppKit.h>
#import <Foundation/Foundation.h>
#import <Vision/Vision.h>

static NSDictionary *Recognize(NSString *path) {
    NSImage *image = [[NSImage alloc] initWithContentsOfFile:path];
    if (image == nil) {
        return @{ @"path": path, @"lines": @[], @"error": @"无法读取页面图像" };
    }

    NSRect rect = NSMakeRect(0, 0, image.size.width, image.size.height);
    CGImageRef cgImage = [image CGImageForProposedRect:&rect context:nil hints:nil];
    if (cgImage == nil) {
        return @{ @"path": path, @"lines": @[], @"error": @"无法转换页面图像" };
    }

    VNRecognizeTextRequest *request = [[VNRecognizeTextRequest alloc] init];
    request.recognitionLevel = VNRequestTextRecognitionLevelAccurate;
    request.usesLanguageCorrection = YES;
    request.minimumTextHeight = 0.006;

    NSError *languageError = nil;
    NSArray<NSString *> *supported = [VNRecognizeTextRequest
        supportedRecognitionLanguagesForTextRecognitionLevel:VNRequestTextRecognitionLevelAccurate
        revision:request.revision
        error:&languageError];
    NSMutableArray<NSString *> *preferred = [NSMutableArray array];
    for (NSString *language in @[ @"zh-Hans", @"zh-Hant", @"en-US" ]) {
        if ([supported containsObject:language]) {
            [preferred addObject:language];
        }
    }
    if (preferred.count > 0) {
        request.recognitionLanguages = preferred;
    }

    VNImageRequestHandler *handler = [[VNImageRequestHandler alloc] initWithCGImage:cgImage options:@{}];
    NSError *recognitionError = nil;
    if (![handler performRequests:@[ request ] error:&recognitionError]) {
        return @{
            @"path": path,
            @"lines": @[],
            @"error": recognitionError.localizedDescription ?: @"Vision OCR 失败"
        };
    }

    NSArray<VNRecognizedTextObservation *> *observations = [request.results sortedArrayUsingComparator:
        ^NSComparisonResult(VNRecognizedTextObservation *left, VNRecognizedTextObservation *right) {
            CGFloat verticalDifference = fabs(CGRectGetMidY(left.boundingBox) - CGRectGetMidY(right.boundingBox));
            if (verticalDifference > 0.012) {
                return CGRectGetMidY(left.boundingBox) > CGRectGetMidY(right.boundingBox)
                    ? NSOrderedAscending
                    : NSOrderedDescending;
            }
            if (CGRectGetMinX(left.boundingBox) < CGRectGetMinX(right.boundingBox)) {
                return NSOrderedAscending;
            }
            return NSOrderedDescending;
        }];

    NSMutableArray<NSDictionary *> *lines = [NSMutableArray array];
    for (VNRecognizedTextObservation *observation in observations) {
        VNRecognizedText *candidate = [observation topCandidates:1].firstObject;
        if (candidate == nil || candidate.string.length == 0) {
            continue;
        }
        [lines addObject:@{
            @"text": candidate.string,
            @"confidence": @(candidate.confidence),
            @"x": @(CGRectGetMinX(observation.boundingBox)),
            @"y": @(CGRectGetMinY(observation.boundingBox))
        }];
    }
    return @{ @"path": path, @"lines": lines, @"error": [NSNull null] };
}

int main(int argc, const char *argv[]) {
    @autoreleasepool {
        for (int index = 1; index < argc; index++) {
            NSString *path = [NSString stringWithUTF8String:argv[index]];
            NSDictionary *result = Recognize(path);
            NSData *data = [NSJSONSerialization dataWithJSONObject:result options:0 error:nil];
            NSString *json = [[NSString alloc] initWithData:data encoding:NSUTF8StringEncoding];
            printf("%s\n", json.UTF8String);
        }
    }
    return 0;
}
