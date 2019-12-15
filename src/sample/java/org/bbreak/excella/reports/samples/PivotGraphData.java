/*-
 * #%L
 * excella-reports
 * %%
 * Copyright (C) 2009 - 2019 bBreak Systems and contributors
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

package org.bbreak.excella.reports.samples;

import java.math.BigDecimal;
import java.util.Date;

/**
 * ピボットサンプル　データクラス
 *
 * @since 1.0
 */
public class PivotGraphData {

    /** 販売日 */
    private Date salesDate;

    /** 商品名 */
    private String productName;

    /** 単価 */
    private BigDecimal unitPrice;

    /** 数量 */
    private BigDecimal quantity;

    /** 金額 */
    private BigDecimal price;

    /** 担当者 */
    private String salesPerson;

    public Date getSalesDate() {
        return salesDate;
    }

    public void setSalesDate( Date salesDate) {
        this.salesDate = salesDate;
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName( String productName) {
        this.productName = productName;
    }

    public BigDecimal getUnitPrice() {
        return unitPrice;
    }

    public void setUnitPrice( BigDecimal unitPrice) {
        this.unitPrice = unitPrice;
    }

    public BigDecimal getQuantity() {
        return quantity;
    }

    public void setQuantity( BigDecimal quantity) {
        this.quantity = quantity;
    }

    public BigDecimal getPrice() {
        return price;
    }

    public void setPrice( BigDecimal price) {
        this.price = price;
    }

    public String getSalesPerson() {
        return salesPerson;
    }

    public void setSalesPerson( String salesPerson) {
        this.salesPerson = salesPerson;
    }
    
}
