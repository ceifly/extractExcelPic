<?php

#[AllowDynamicProperties] class extractExcelPic
{

    private string $file_path;
    private string $zip_name;
    private array $allow_excel = ['xlsx'];
    protected string $extract_path;
    public string $prefix = "ext";
    public string $origin_file;
    public string $xml_path = "xl/drawings/drawing{{number}}.xml";
    public string $xml_rel_path = "xl/drawings/_rels/drawing{{number}}.xml.rels";
    public string $media_path = "xl/media";
    public array $path_info;
    private array $pic_lists;
    private array $file_map;
    private string|int $max_drawing;
    //private array $xml_pic_map;

    public function __construct(string $origin_file, string $file_path = '')
    {
        $this->origin_file = $origin_file;
        $this->file_path = $file_path;
    }

    public function __set(string $name, $value): void
    {
        if ($name === 'origin_file') {
            $this->origin_file = $value;
        }
    }

    public function run(): array
    {
        $response = ['code' => 0, 'message' => '', 'data' => []];

        try {
            $this->getPathInfo();
            $this->copyFile();
            $this->extractZip($this->file_path . $this->zip_name, ['xl/media', 'xl/drawings/drawing', 'xl/drawings/_rels']);
            $xml_res_array = $this->getXmlContent();
            foreach ($xml_res_array as $xml_key => $xml_map) {

                $xml_array = $this->parseXmlContent($xml_map);
                $xml_pic[$xml_key] = $this->parseXmlPicContent($xml_map);
                $response['data'][$xml_key] = $this->parseRow($xml_key , $xml_array , $xml_pic);
            }
            $this->removeZip();
            return $response;
        } catch (Exception $exception) {
            $response['code'] = $exception->getCode();
            $response['message'] = $exception->getMessage();
        }

        return $response;
    }

    /**
     * @throws Exception
     */
    public function copyFile(): void
    {
        if (!file_exists($this->origin_file)) {
            throw new Exception("原始文件不存在", "401");
        }

        if (empty($this->path_info['dirname'])) {
            throw new Exception("上传文件信息有误", "501");
        }

        if (!in_array($this->path_info['extension'], $this->allow_excel)) {
            throw new Exception("文件格式有误", "502");
        }

        $zip_name = $this->makeZipName();

        if (empty($this->file_path)) {
            $this->file_path = $this->path_info['dirname'] . "/";
        }

        $res = copy($this->origin_file, $this->file_path . $zip_name);
        if (!$res) {
            throw new Exception("复制文件失败", "503");
        }
    }

    private function makeZipName(): string
    {
        $prefix = $this->prefix;
        $file_name = $prefix . '_' . (time()) . "_" . mt_rand(100, 999);
        return $this->zip_name = $file_name . '.zip';
    }

    public function getPathInfo(): array|string
    {
        $path_info = pathinfo($this->origin_file);
        return $this->path_info = is_array($path_info) ? $path_info : [];
    }

    /**
     * @throws Exception
     */
    protected function extractZip($zipFile = '', $need_files = []): void
    {
        if (!file_exists($zipFile)) {
            throw new Exception("转换zip文件失败", "601");
        }
        //解压
        $path_info = pathinfo($zipFile);
        $this->extract_path = $this->file_path . $path_info['filename'] . "/";

        $zip = new ZipArchive();

        $res = $zip->open($zipFile, ZipArchive::RDONLY);
        if ($res !== true) {
            throw new Exception("zip文件不存在", "602");
        }
        $file_map = [];
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $zip_entry_name = $zip->getNameIndex($i);
            $complete_path = $this->extract_path . $zip_entry_name;
            $save = false;

            if (str_starts_with($zip_entry_name, "xl/drawings/drawing")) {
                preg_match("/drawing(\d+)\.xml/", $zip_entry_name, $mat);
                $this->max_drawing = ++$mat[1] ?? 2;
            }

            foreach ($need_files as $need_file) {
                if (str_starts_with($zip_entry_name, $need_file)) {
                    $save = true;
                }
            }
            if (!str_ends_with($complete_path, '/') && !file_exists($complete_path) && $save) {
                $path_info = pathinfo($complete_path);
                $tmp = '';
                foreach (explode('/', $path_info['dirname']) as $k) {
                    $tmp .= $k . '/';
                    if (!is_dir($tmp)) {
                        @mkdir($tmp);
                    }
                }
                $res = @file_put_contents($complete_path, $zip->getStream($zip_entry_name));
                if ($res === false) {
                    throw new Exception("zip文件解压失败", "602");
                }
                if (str_starts_with($zip_entry_name, $this->media_path)) {
                    $file_map[$path_info['filename']] = $this->setPathInfo($path_info, $complete_path);
                }
            }
        }
        $this->file_map = $file_map;
        /*$res = $zip->extractTo($this->extract_path);
        if (!$res) {
            throw new Exception("zip文件解压失败", "602");
        }*/
        $zip->close();
    }

    /**
     * @throws Exception
     */
    protected function parseXmlContent($xml_map): array
    {
        try {
            $xml_array = $this->covertToArray($xml_map['xml_string']);
        } catch (Exception $exception) {
            throw new  Exception($exception->getMessage(), $exception->getCode());
        }
        return $xml_array;
    }

    /**
     * @throws Exception
     */
    protected function parseXmlPicContent($xml_map): array
    {
        try {
            $xml_rel_array = $this->covertToArray($xml_map['xml_rel_string']);
        } catch (Exception $exception) {
            throw new  Exception($exception->getMessage(), $exception->getCode());
        }

        if (empty($xml_rel_array['Relationship'])) {
            throw new  Exception("xml图片关系解析失败", "1002");
        }
        $xml_pic_map = [];

        if (isset($xml_rel_array['Relationship']['@attributes'])) {
            $xml_pic_dom[]= $xml_rel_array['Relationship'];
        } else {
            $xml_pic_dom = $xml_rel_array['Relationship'];
        }

        foreach ($xml_pic_dom as $item) {
            preg_match("/(image\d+)/", $item['@attributes']['Target'] ?? "" , $mat);

            if (!empty($mat['1'])) {
                $xml_pic_map[$item['@attributes']['Id']] = $mat['1'];
            }
        }

        //$this->xml_pic_map[$xml_key] = $xml_pic_map;
        return $xml_pic_map;
    }

    /**
     * @throws Exception
     */
    protected function getXmlContent(): array
    {
        $max = $this->max_drawing ?? 2;
        $xml_content_map = [];
        for ($i = 1; $i < $max; $i++) {
            $xml_path = $this->extract_path . str_replace("{{number}}", $i, $this->xml_path);
            if (!file_exists($xml_path)) {
                throw new  Exception("xml文件不存在,请检查上传文件", "701");
            }

            $xml_string = file_get_contents($xml_path);
            if (!$xml_string) {
                throw new  Exception("xml获取失败", "702");
            }

            $xml_rel_path = $this->extract_path . str_replace("{{number}}", $i, $this->xml_rel_path);
            if (!file_exists($xml_rel_path)) {
                throw new  Exception("xml.rel文件不存在,请检查上传文件", "701");
            }

            $xml_rel_string = file_get_contents($xml_rel_path);
            if (!$xml_rel_string) {
                throw new  Exception("xml.rel获取失败", "702");
            }
            $xml_content_map[$i]['xml_string'] = $xml_string;
            $xml_content_map[$i]['xml_rel_string'] = $xml_rel_string;
        }
        return $xml_content_map;
    }

    /**
     * @throws Exception
     */
    protected function covertToArray(string $xml): array
    {
        $res = assert(class_exists('\DOMDocument'));
        if ($res === false) {
            throw new  Exception("DOMDocument未安装", "801");
        }

        $doc = new DOMDocument();
        $doc->loadXML($xml);
        $root = $doc->documentElement;
        $output = (array)$this->domNodeToArray($root);

        return $output ?? [];
    }

    protected function domNodeToArray($node): array|string
    {
        $output = [];

        switch ($node->nodeType) {
            case 4: // XML_CDATA_SECTION_NODE
            case 3: // XML_TEXT_NODE
                $output = trim($node->textContent);
                break;
            case 1: // XML_ELEMENT_NODE
                for ($i = 0, $m = $node->childNodes->length; $i < $m; $i++) {
                    $child = $node->childNodes->item($i);
                    $v = $this->domNodeToArray($child);
                    if (isset($child->tagName)) {
                        $t = $child->tagName;
                        if (!isset($output[$t])) {
                            $output[$t] = [];
                        }
                        if (empty($v)) {
                            $v = '';
                        }
                        $output[$t][] = $v;
                    } elseif ($v) {
                        $output = (string)$v;
                    }
                }

                if ($node->attributes->length && !is_array($output)) { // has attributes but isn't an array
                    $output = ['@content' => $output]; // change output into an array.
                }
                if (is_array($output)) {
                    if ($node->attributes->length) {
                        $a = [];
                        foreach ($node->attributes as $attrName => $attrNode) {
                            $a[$attrName] = (string)$attrNode->value;
                        }
                        $output['@attributes'] = $a;
                    }
                    foreach ($output as $t => $v) {
                        if ($t !== '@attributes' && is_array($v) && count($v) === 1) {
                            $output[$t] = $v[0];
                        }
                    }
                }
                break;
        }

        return $output;
    }

    /**
     * @throws Exception
     */
    public function parseRow($xml_key , $xml_array , $xml_pic_map): array
    {
        $pics = [];
        if (!empty($xml_array['xdr:twoCellAnchor'])) {
            $temp_dom_arr = $xml_array['xdr:twoCellAnchor'];
        } else {
            $temp_dom_arr = $xml_array['xdr:oneCellAnchor'] ?? [];
        }

        if (empty($temp_dom_arr)) {
            throw new  Exception("xml解析失败", "903");
        }

        if (isset($temp_dom_arr['xdr:from'])) {
            $dom_arr[] = $temp_dom_arr;
        } else {
            $dom_arr = $temp_dom_arr;
        }

        if (empty($dom_arr)) {
            throw new  Exception("xml解析失败", "901");
        }

        $pic_lists = [];
        foreach ($dom_arr as $dom) {

            $row = $dom['xdr:from']['xdr:row'] ?? 0;//行
            $col = $dom['xdr:from']['xdr:col'] ?? 0;//列
            if (!$row && !$col) {
                throw new  Exception("xml解析失败", "902");
            }
            $pic = $dom['xdr:pic']['xdr:blipFill']['a:blip']['@attributes']['embed'] ?? "";
            if ($pic) {

                $temp_name = str_replace("rId", "image", $pic);
                $pics[$col][$row][] = $temp_name;
                $pic_lists[] = [
                    'sheet' => $xml_key,
                    'name' => $temp_name,
                    'row' => $row,
                    'col' => $col,
                    'path_info' => $this->file_map[$xml_pic_map[$xml_key][$pic]??""] ?? [],

                ];
            }
        }
        $this->pic_lists[$xml_key] = $pic_lists;
        return $pics;
    }

    private function removeZip(): void
    {
        if (file_exists($this->file_path . $this->zip_name)) {
            @unlink($this->file_path . $this->zip_name);
        }
    }

    public function getPicList(): array
    {
        return $this->pic_lists ?? [];
    }

    protected function setPathInfo($path_info, $complete_path): array
    {
        $file_map = $path_info;
        $file_map['complete_path'] = $complete_path;
        $file_map['media_path'] = $this->media_path;
        if (file_exists($complete_path)) {
            $file_map['md5'] = md5($complete_path);
            //$file_map['base64_encode'] = base64_encode(file_get_contents($complete_path));
        } else {
            $file_map['md5'] = '';
            $file_map['base64_encode'] = '';
        }
        return $file_map;
    }

}

$origin_name1 = 'D:/SubjectForTest/file/package_HYmodel_test_has_picture.xlsx';
$origin_name2 = 'D:/SubjectForTest/file/package_HYmodel_test_1_picture.xlsx';
$pic = new extractExcelPic('', "D:/SubjectForTest/file/123/");
/*$pic->origin_file = $origin_name1;
$res = $pic->run();*/
$pic->origin_file = $origin_name2;
$pic->run();
$res = $pic->getPicList();
var_dump($res);
