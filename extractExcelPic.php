<?php

#[AllowDynamicProperties] class extractExcelPic
{

    private string $file_path;
    private string $zip_name;
    private array $allow_excel = ['xlsx'];
    protected string $extract_path;
    public string $prefix = "ext";
    public string $origin_file;
    public string $xml_path = "/xl/drawings/drawing1.xml";
    public array $path_info;

    public function __construct(string $origin_file, string $file_path = '')
    {
        $this->origin_file = $origin_file;
        $this->file_path = $file_path;
        $this->getPathInfo();
    }

    public function run(): array
    {
        $response = ['code' => 0, 'message' => '', 'data' => []];

        try {
            $this->copyFile();
            $this->extractZip($this->file_path . $this->zip_name, ['xl/media', 'xl/drawings/drawing1.xml']);
            $xml_string = $this->getXmlContent();
            $xml_array = $this->covertToArray($xml_string);
            $response['data'] = $this->parseRow($xml_array);
            $this->removeZip();
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

        for ($i = 0; $i < $zip->numFiles; $i++) {
            $zip_entry_name = $zip->getNameIndex($i);
            $complete_path = $this->extract_path . $zip_entry_name;
            $save = false;
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
                        @mkdir($tmp, 0777);
                    }
                }
                $res = @file_put_contents($complete_path, $zip->getStream($zip_entry_name));
                if ($res === false) {
                    throw new Exception("zip文件解压失败", "602");
                }
            }
        }
        /*$res = $zip->extractTo($this->extract_path);
        if (!$res) {
            throw new Exception("zip文件解压失败", "602");
        }*/
        $zip->close();
    }

    /**
     * @throws Exception
     */
    protected function getXmlContent(): string
    {
        $xml_path = $this->extract_path . $this->xml_path;
        if (!file_exists($xml_path)) {
            throw new  Exception("xml文件不存在,请检查上传文件", "701");
        }

        $xml_string = file_get_contents($xml_path);
        if (!$xml_string) {
            throw new  Exception("xml获取失败", "702");
        }

        return $xml_string;
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
    public function parseRow($xml_array): array
    {
        $pics = [];
        if (!isset($xml_array['xdr:twoCellAnchor'])) {
            throw new  Exception("xml解析失败", "901");
        }
        foreach ($xml_array['xdr:twoCellAnchor'] as $dom) {
            $row = $dom['xdr:from']['xdr:row'] ?? 0;
            if (!$row) {
                throw new  Exception("xml解析失败", "902");
            }
            $pic = $dom['xdr:pic']['xdr:blipFill']['a:blip']['@attributes']['embed'] ?? "";
            if ($pic) {
                $pics[$row][] = str_replace("rId", "image", $pic);
            }
        }
        return $pics;
    }

    private function removeZip(): void
    {
        if (file_exists($this->file_path . $this->zip_name)) {
            @unlink($this->file_path . $this->zip_name);
        }
    }

}

$origin_name = 'D:/SubjectForTest/file/package_HYmodel_test_has_picture.xlsx';
$pic = new extractExcelPic($origin_name , "D:/SubjectForTest/file/123/");
$res = $pic->run();
var_dump($res);
