from wasmtime.loader import store
import wasmtime.loader
import egg

BUFFER = bytearray()

def typstArgs(*A):
    for i in A:
        for j in i:
            BUFFER.append(j)
    return [len(i) for i in A]

# Define the required imports as functions
def wasm_minimal_protocol_write_args_to_buffer(arg1_ptr):
    for idx,i in enumerate(list(BUFFER)):
        egg.memory.data_ptr(store)[arg1_ptr+idx]=i


def wasm_minimal_protocol_send_result_to_host(arg1_ptr, arg2_len):
    data = egg.memory.data_ptr(store)[arg1_ptr:arg1_ptr+arg2_len]
    print(data)
    print(bytes(data).decode('utf-8'))
    pass
