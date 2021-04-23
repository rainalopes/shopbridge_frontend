import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class ProductService {

  constructor(private http: HttpClient) { }
  getProducts() {
    return this.http.get<any>('assets/product.json')
    .toPromise()
    .then(res => res.data)
    .then(data => { return data; });
}
}
