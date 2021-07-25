import { Component, OnInit } from '@angular/core';
import { ProductService } from './../product-service.service';
import { ConfirmationService } from 'primeng/api';
import { MessageService } from 'primeng/api';
import * as Excel from "exceljs";
import * as fs from 'file-saver';
@Component({
  selector: 'app-product-management',
  templateUrl: './product-management.component.html',
  styleUrls: ['./product-management.component.css']
})
export class ProductManagementComponent implements OnInit {

productDialog: boolean;
products: any[];
product: any;
selectedProducts: any[];
submitted: boolean;
title: string;
header:any;
exceldata:any[];
    dropdownlist: any;
  constructor(private productService: ProductService, private messageService: MessageService, private confirmationService: ConfirmationService) { }
  
  ngOnInit() {
      this.productService.getProducts().then(data => {
          this.products = data;
          this.title = 'Sales Report';
          this.header = ["Name", "Price", "Category"]
          this.exceldata = this.products.map((element)=>{
              return [element.name,element.price,element.category];
          });
    });
    this.productService.getCategories().then(
        data=>{
            this.dropdownlist = [];
            this.dropdownlist = data.map((item)=>{
                return item.category;
            });
        }
    );
  }

  generateExcel(){
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet('Products Data');
    let titleRow = worksheet.addRow([this.title]);
    // Set font, size and style in title row.
    titleRow.font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };
    // Blank Row
    worksheet.addRow([]);
    //Add Header Row
    let headerRow = worksheet.addRow(this.header);

    // Cell Style : Fill and Border
    headerRow.eachCell((cell, number) => {
    cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' },
        bgColor: { argb: 'FF0000FF' }
    }
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    });
    // Add Data and Conditional Formatting
    console.log(this.exceldata);
    this.exceldata.forEach(d => {
        let row = worksheet.addRow(d);
    }
    
    );
    workbook.xlsx.writeBuffer().then((data) => {
        let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        fs.saveAs(blob, 'Report.xlsx');
  });

  }
  generateTemplate(){
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet('Products Data');
    let titleRow = worksheet.addRow([this.title]);
    // Set font, size and style in title row.
    titleRow.font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };
    // Blank Row
    worksheet.addRow([]);
    //Add Header Row
    let headerRow = worksheet.addRow(this.header);

    // Cell Style : Fill and Border
    headerRow.eachCell((cell, number) => {
    cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' },
        bgColor: { argb: 'FF0000FF' }
    }
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    });
    // Add Data and Conditional Formatting
    let joineddropdownlist = "\""+this.dropdownlist.join(',')+"\"";
    console.log(joineddropdownlist);
    for(let i=4;i<100;i++){
        worksheet.getCell('C'+i).dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [joineddropdownlist]//'"One,Two,Three,Four"'
          };
    }
    workbook.xlsx.writeBuffer().then((data) => {
        let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        fs.saveAs(blob, 'Template.xlsx');
  });
  }
  openNew() {
      this.product = {};
      this.submitted = false;
      this.productDialog = true;
  }

  deleteSelectedProducts() {
      this.confirmationService.confirm({
          message: 'Are you sure you want to delete the selected products?',
          header: 'Confirm',
          icon: 'pi pi-exclamation-triangle',
          accept: () => {
              this.products = this.products.filter(val => !this.selectedProducts.includes(val));
              this.selectedProducts = null;
              this.messageService.add({severity:'success', summary: 'Successful', detail: 'Products Deleted', life: 3000});
          }
      });
  }

  editProduct(product) {
      this.product = {...product};
      this.productDialog = true;
  }

  deleteProduct(product) {
      this.confirmationService.confirm({
          message: 'Are you sure you want to delete ' + product.name + '?',
          header: 'Confirm',
          icon: 'pi pi-exclamation-triangle',
          accept: () => {
              this.products = this.products.filter(val => val.id !== product.id);
              this.product = {};
              this.messageService.add({severity:'success', summary: 'Successful', detail: 'Product Deleted', life: 3000});
          }
      });
  }

  hideDialog() {
      this.productDialog = false;
      this.submitted = false;
  }
  
  saveProduct() {
      this.submitted = true;

      if (this.product.name.trim()) {
          if (this.product.id) {
              this.products[this.findIndexById(this.product.id)] = this.product;                
              this.messageService.add({severity:'success', summary: 'Successful', detail: 'Product Updated', life: 3000});
          }
          else {
              this.product.id = this.createId();
              this.products.push(this.product);
              this.messageService.add({severity:'success', summary: 'Successful', detail: 'Product Created', life: 3000});
          }

          this.products = [...this.products];
          this.productDialog = false;
          this.product = {};
      }
  }

  findIndexById(id: string): number {
      let index = -1;
      for (let i = 0; i < this.products.length; i++) {
          if (this.products[i].id === id) {
              index = i;
              break;
          }
      }

      return index;
  }

  createId(): string {
      let id = '';
      var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
      for ( var i = 0; i < 5; i++ ) {
          id += chars.charAt(Math.floor(Math.random() * chars.length));
      }
      return id;
  }

}
