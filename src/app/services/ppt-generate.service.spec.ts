import { TestBed } from '@angular/core/testing';

import { PptGenerateService } from './ppt-generate.service';

describe('PptGenerateService', () => {
  let service: PptGenerateService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(PptGenerateService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
