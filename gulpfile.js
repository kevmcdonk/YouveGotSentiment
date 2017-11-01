'use strict';

var gulp = require('gulp');
var deployCdn = require('gulp-deploy-azure-cdn');
var gutil = require('gulp-util');
var fs = require('fs');
var gulpIgnore = require('gulp-ignore');
var path = require('path');

var accountName = ''; // Azure Blob Storage account name
var accountKey = ''; // Azure Blob Storage account key

gulp.task('publishAzure', function () {
    return gulp.src(['dist/**'], {
        base: 'dist' // optional, the base directory in which the file is located. The relative path of file to this directory is used as the destination path
    }).pipe(deployCdn({
        containerName: 'youvegotsentimentprod', // container name in blob
        serviceOptions: [accountName, accountKey], // custom arguments to azure.createBlobService
        //folder: '1.2.35-b27', // path within container
        //zip: true, // gzip files if they become smaller after zipping, content-encoding header will change if file is zipped
        deleteExistingBlobs: true, // true means recursively deleting anything under folder
        concurrentUploadThreads: 10, // number of concurrent uploads, choose best for your network condition
        metadata: {
            cacheControl: 'public, max-age=31530000', // cache in browser
            cacheControlHeader: 'public, max-age=31530000' // cache in azure CDN. As this data does not change, we set it to 1 year
        },
        testRun: false // test run - means no blobs will be actually deleted or uploaded, see log messages for details
    })).on('error', gutil.log);
});
