import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import * as AWS from 'aws-sdk';
// import { S3 } from '@aws-sdk/client-s3';
import { Observable } from 'rxjs';
import { Utils } from 'tslint';
import * as S3 from 'aws-sdk/clients/s3';

@Injectable({
    providedIn: 'root',
})
export class S3UtilService {

    bucket_Name = "venus42";


    uploadfile(file: any, FOLDER: any, filename: any) {
        let accessKeyId = 'AKIA2VMXRTPE342SFXPY';
        let secretAccessKey = '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';
        const bucket = new S3(
            {
                accessKeyId: accessKeyId,
                secretAccessKey: secretAccessKey,
                region: 'eu-central-1'
            }
        );

        const params = {
            Bucket: this.bucket_Name,
            Key: FOLDER + filename,
            Body: file
        };

        const options = {
            // Part Size of 10mb
            partSize: 5 * 1024 * 1024,
            queueSize: 1,
            // Give the owner of the bucket full control
            ACL: '908405280a6bc0be58c466f59004b7d77e2f42a68e351b9c99de4a199cc04fc2'
        };

        return new Promise(function (resolve, reject) {
            bucket.upload(params, options, function (error: any, result: any) {
                if (error) {
                    console.log(error);
                    reject(false);
                }

                resolve(true);
            });
        });


    }

    getFileFromS3(file_path: any) {
        let accessKeyId = 'AKIA2VMXRTPE342SFXPY';
        let secretAccessKey = '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';
        const options = {
            accessKeyId: accessKeyId,
            secretAccessKey: secretAccessKey,
            region: 'eu-central-1',
        }

        var s3 = new S3(options);
        const url = s3.getSignedUrl('getObject', {
            Bucket: this.bucket_Name,
            Key: file_path,
        })
        return url;
    }
}
