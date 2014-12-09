// Helper for converting SafeArray to array of objects.
package ole

import (
	"fmt"
	"reflect"
	"unsafe"
)

type SafeArrayConversion struct {
	Array *SafeArray
}

func (sac *SafeArrayConversion) ToStringArray() (strings []string) {
	totalElements, _ := sac.TotalElements(0)
	strings = make([]string, totalElements)

	for i := int64(0); i < totalElements; i++ {
		strings[int32(i)], _ = safeArrayGetElementString(sac.Array, i)
	}

	return
}

func (sac *SafeArrayConversion) ToByteArray() (bytes []byte) {
	totalElements, _ := sac.TotalElements(0)
	bytes = make([]byte, totalElements)

	for i := int64(0); i < totalElements; i++ {
		ptr, _ := safeArrayGetElement(sac.Array, i)
		bytes[int32(i)] = *(*byte)(unsafe.Pointer(&ptr))
	}

	return
}

func (sac *SafeArrayConversion) ToVariantArray(results *[]VARIANT) (err error) {
	dv := reflect.ValueOf(results)
	totalElements, err := sac.TotalElements(0)
	if err != nil {
		fmt.Printf("ToVariantArray TotalElements err %v\n", err)
		return err
	}
	fmt.Printf("ToVariantArray results len=%v cap=%v\n", len(*results), cap(*results))
	//*results = make([]VARIANT, 0, totalElements)
	dv.Elem().Set(reflect.MakeSlice(dv.Elem().Type(), int(totalElements), int(totalElements)))
	fmt.Printf("ToVariantArray results len=%v cap=%v\n", len(*results), cap(*results))
	for i := int64(0); i < totalElements; i++ {
		ptr, err := safeArrayGetElement(sac.Array, int64(i))
		if err != nil {
			fmt.Printf("ToVariantArray safeArrayGetElement err %v\n", err)
			return err
		}
		//(*results)[i] = *(*VARIANT)(unsafe.Pointer(&ptr))
		ptrUnsafe := unsafe.Pointer(&ptr)
		ptruintptr := uintptr(ptrUnsafe)
		ptrint := int(ptruintptr)
		fmt.Printf("ToVariantArray ptr=%p ptrUnsafe=%p ptruintptr=%#v ptrint=%#v\n", ptr, ptrUnsafe, ptruintptr, ptrint)
		var v *VARIANT = &VARIANT{VT_DISPATCH | VT_BYREF, 0, 0, 0, int(uintptr(unsafe.Pointer(&ptr))), 0}
		(*results)[i] = *v
		//(*results)[i] = NewVariant(VT_VARIANT, uint64(uintptr(unsafe.Pointer(&ptr))))
		//(*results)[i] = NewVariant(VT_VARIANT|VT_BYREF, uint64(uintptr(unsafe.Pointer(&ptr))))
		fmt.Printf("ToVariantArray loop i=%v ptr=%p result[i]=%v\n", i, ptr, (*results)[i])
	}
	fmt.Printf("ToVariantArray finished results: %v\n", *results)
	return
}

func (sac *SafeArrayConversion) GetType() (varType uint16, err error) {
	return safeArrayGetVartype(sac.Array)
}

func (sac *SafeArrayConversion) GetDimensions() (dimensions *uint32, err error) {
	return safeArrayGetDim(sac.Array)
}

func (sac *SafeArrayConversion) GetSize() (length *uint32, err error) {
	return safeArrayGetElementSize(sac.Array)
}

func (sac *SafeArrayConversion) TotalElements(index uint32) (totalElements int64, err error) {
	if index < 1 {
		index = 1
	}

	// Get array bounds
	var LowerBounds int64
	var UpperBounds int64

	LowerBounds, err = safeArrayGetLBound(sac.Array, index)
	if err != nil {
		return
	}

	UpperBounds, err = safeArrayGetUBound(sac.Array, index)
	if err != nil {
		return
	}

	totalElements = UpperBounds - LowerBounds + 1
	return
}

// Release Safe Array memory
func (sac *SafeArrayConversion) Release() {
	safeArrayDestroy(sac.Array)
}
