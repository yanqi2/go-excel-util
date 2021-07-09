package excelizeUtil

import (
	"bytes"
	"crypto/tls"
	"encoding/json"
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"image"
	"io"
	"io/ioutil"
	"net/http"
	"os"
	"path"
	"path/filepath"
	"strings"
)

const (
	defaultCellWidth  = 35
	defaultCellHeight = 180
)

var supportImageTypes = map[string]string{".gif": ".gif", ".jpg": ".jpeg", ".jpeg": ".jpeg", ".png": ".png", ".tif": ".tiff", ".tiff": ".tiff"}

type formatPicture struct {
	FPrintsWithSheet bool    `json:"print_obj"`
	FLocksWithSheet  bool    `json:"locked"`
	NoChangeAspect   bool    `json:"lock_aspect_ratio"`
	Autofit          bool    `json:"autofit"`
	OffsetX          int     `json:"x_offset"`
	OffsetY          int     `json:"y_offset"`
	XScale           float64 `json:"x_scale"`
	YScale           float64 `json:"y_scale"`
	Hyperlink        string  `json:"hyperlink"`
	HyperlinkType    string  `json:"hyperlink_type"`
	Positioning      string  `json:"positioning"`
}

func AddPicture(f *excelize.File, sheet, cell, picture, format string) error {
	var err error
	// 检查 picture 格式
	if _, err = os.Stat(picture); os.IsNotExist(err) {
		return err
	}
	ext, ok := supportImageTypes[path.Ext(picture)]
	if !ok {
		return errors.New("unsupported image extension")
	}

	var file []byte
	var img image.Config
	_, name := filepath.Split(picture)
	width := &img.Width
	height := &img.Height

	if strings.HasPrefix(picture, "http") {
		// 获取网络图片
		tr := &http.Transport{
			TLSClientConfig: &tls.Config{InsecureSkipVerify: true},
		}
		client := &http.Client{Transport: tr}
		resp, err := client.Get(picture)
		if err != nil {
			fmt.Println("client.Get err  :", err)
			return err
		}
		var buf bytes.Buffer
		tee := io.TeeReader(resp.Body, &buf)
		file, _ = ioutil.ReadAll(tee)
		img, _, _ = image.DecodeConfig(&buf)
		width = &img.Width
	} else {
		// 获取本地图片
		file, _ = ioutil.ReadFile(picture)
		img, _, err = image.DecodeConfig(bytes.NewReader(file))
	}

	var scale float64
	if *height > *width {
		// 以 height 为基准
		scale = 1.23 * defaultCellHeight / float64(*height)
		if err = f.SetColWidth(sheet, getColName(colIndex), getColName(colIndex), float64(defaultCellHeight**width / *height)/5.2); err != nil {
			return err
		}
		if err = f.SetRowHeight(sheet, int(rowIndex+1), defaultCellHeight); err != nil {
			return err
		}
	} else {
		scale = 6.5 * defaultCellWidth / float64(*width)
		if err = f.SetColWidth(sheet, getColName(colIndex), getColName(colIndex), defaultCellWidth); err != nil {
			return err
		}
		if err = f.SetRowHeight(sheet, int(rowIndex+1), 5.8*float64(defaultCellWidth**height / *width)); err != nil {
			return err
		}
	}

	format, err = parseFormatPictureSet(format, scale, picture)
	if err != nil {
		return err
	}

	return f.AddPictureFromBytes(sheet, cell, format, name, ext, file)
}

func parseFormatPictureSet(formatSet string, defaultScale float64, picture string) (formatSetParsed string, err error) {
	format := formatPicture{
		FPrintsWithSheet: true,
		FLocksWithSheet:  true,
		NoChangeAspect:   true,
		Autofit:          false,
		OffsetX:          0,
		OffsetY:          0,
		XScale:           defaultScale,
		YScale:           defaultScale,
	}
	if strings.HasPrefix(picture, "http") {
		//format.Hyperlink = picture
		//format.HyperlinkType = "External"
	}
	if err = json.Unmarshal(parseFormatSet(formatSet), &format); err != nil {
		return "", err
	}
	data, err := json.Marshal(format)
	return string(data), err
}

func parseFormatSet(formatSet string) []byte {
	if formatSet != "" {
		return []byte(formatSet)
	}
	return []byte("{}")
}
