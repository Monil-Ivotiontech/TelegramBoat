import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable, forkJoin } from 'rxjs';
import { environment } from 'src/environments/environment';
import { EndPoints, ApiMethod } from '../const';

@Injectable({
  providedIn: 'root'
})
export class HttpService {

  constructor(private http: HttpClient) { }


  frontendRequestCall(api: EndPoints, method: ApiMethod, data?: any): Observable<any> {
    switch (method) {

      case ApiMethod.GET:
        return this.http.get(`${environment.backend_api_url}${api}`)

      case ApiMethod.POST:
        return this.http.post(`${environment.backend_api_url}${api}`, data)


      case ApiMethod.PUT:
        return this.http.put(`${environment.backend_api_url}${api}`, data)


      case ApiMethod.DELETE:
        return this.http.delete(`${environment.backend_api_url}${api}`)

    }
  }

  frontendRequestMutliCall(requestList: Array<any>): Observable<any[]> {
    let forkJoinRequests: Array<any> = [];
    requestList.forEach(item => {
      forkJoinRequests.push(this.frontendRequestCall(item.api, item.method, item.data));
    })
    return forkJoin(forkJoinRequests);
  }




}
