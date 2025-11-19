import { TestBed } from '@angular/core/testing';

import { S3ClientService } from './s3-client.service';

describe('S3ClientService', () => {
  let service: S3ClientService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(S3ClientService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
