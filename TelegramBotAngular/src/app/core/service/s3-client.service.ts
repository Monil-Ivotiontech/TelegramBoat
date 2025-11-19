import { Injectable } from '@angular/core';
import { S3 } from "@aws-sdk/client-s3";

@Injectable({
  providedIn: 'root'
})
export class S3ClientService {

  constructor() { }

  s3_client_method() {
    let s3Client = new S3({
      forcePathStyle: false, // Configures to use subdomain/virtual calling format.
      endpoint: "https://sgp1.digitaloceanspaces.com",
      region: "us-east-1",
      credentials: {
        accessKeyId: "DO003HDFQD2TNZY7AJBY",
        secretAccessKey: "9p1gUP4uiDPUj81E1aq+jB6h2j2r9sV35QY4cF8hKOo"
      }
    });

    return s3Client;
  }
}

