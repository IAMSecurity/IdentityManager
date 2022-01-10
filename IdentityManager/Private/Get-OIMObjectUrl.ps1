Function Get-OIMObjectUrl {
    Param($object)
    ForEach($item in $object.xObjectKey){
					
        $xmlXObjectKey = 	[xml] $item
        "$Script:BaseURI/api/entity/$($xmlXObjectKey.key.T))/$($xmlXObjectKey.key.P)"
    }
}