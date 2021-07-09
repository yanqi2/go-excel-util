package excelizeUtilV2

import (
	"reflect"
	"testing"
)

type Order struct {
	Id      int64         `name:"订单号" type:"string"`
	Buyer   string        `name:"买家"`
	Seller  string        `name:"卖家"`
	Items   []*Goods      `name:"商品信息" merge:"vertical"`
	Status  int           `name:"订单状态"`
	Image   []string      `name:"快递图片" merge:"horizontal"`
	Express *OrderExpress `name:"快递"`
}

type Goods struct {
	Id         int64   `name:"商品编号" type:"string"`
	Name       string  `name:"商品名"`
	Price      float64 `name:"商品单价"`
	Count      int     `name:"商品个数"`
	TotalPrice float64 `name:"总价格"`
}

type OrderExpress struct {
	Id      string `name:"快递编号"`
	Content string `name:"物流信息"`
}

func TestGenExcel(t *testing.T) {
	file := NewFile()
	orders := GetOrders()
	meta := reflect.TypeOf(Order{})
	file.SetData("订单", meta, ConvertExcelData(orders))
}

func GetOrders() []*Order {
	return []*Order{
		{
			Id:     10000,
			Buyer:  "张三",
			Seller: "李宁官方旗舰店",
			Items: []*Goods{
				{
					Id:         30001,
					Name:       "鞋子",
					Price:      399,
					Count:      2,
					TotalPrice: 798,
				}, {
					Id:         30003,
					Name:       "帽子",
					Price:      99,
					Count:      1,
					TotalPrice: 99,
				},
			},
			Image: []string{
				"http://biu-hk.dwstatic.com/mkcli/20200421/0a4400ef9cd49e5ea0cc89bc50620000.jpg",
				"http://biu-hk.dwstatic.com/mkcli/20200421/0a4400ef1ddc9e5ea0cc966236e90000.jpg",
			},
		}, {
			Id:     10001,
			Buyer:  "李四",
			Seller: "李宁官方旗舰店",
			Items: []*Goods{
				{
					Id:         30001,
					Name:       "鞋子",
					Price:      399,
					Count:      1,
					TotalPrice: 399,
				},
			},
			Image: []string{
				"http://biu-hk.dwstatic.com/mkcli/20200421/0a4400ef7dc89e5ea0cc24b2915a0000.jpg",
				"http://biu-hk.dwstatic.com/mkcli/20200421/0a4400ef52bf9e5ea0ccf9d51a170000.jpg",
				"http://biu-hk.dwstatic.com/mkcli/20200421/0a4400efa6a59e5e9fcc34e8b9de0000.jpg",
				"http://biu-hk.dwstatic.com/mkcli/20200421/0a4400efc9b89e5ea0cc9fed5bca0000.jpg",
				"http://biu-hk.dwstatic.com/mkcli/20200421/0a4400efefbe9e5ea0cc10e118770000.jpg",
			},
		}, {
			Id:     10002,
			Buyer:  "王五",
			Seller: "李宁官方旗舰店",
			Items: []*Goods{
				{
					Id:         30001,
					Name:       "鞋子",
					Price:      399,
					Count:      1,
					TotalPrice: 399,
				},
			},
			Image: nil,
		},
	}
}

func ConvertExcelData(orders []*Order) []interface{} {
	data := make([]interface{}, 0, len(orders))
	for _, order := range orders {
		data = append(data, order)
	}
	return data
}
