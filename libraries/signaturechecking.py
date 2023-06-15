from signature_compare import SignatureComparison



def matchSignature(path):
    sg = SignatureComparison(path)
    return sg.compute_similarity(); signature_status


def sign_checking(path):
    try:
        print(type(path))
        print(path)
        path = str(path)
        print(path)
        sg = matchSignature(path)
        print(sg)
        if float(sg) > 0.21:
            print("True")
            return True
        else:
            print("False")
            return False
    except Exception as e:
        print("An error occured at signature checking:", str(e))
        return  'error'
                
# sign_checking("C:/Users/Q0037/Downloads/signature_verification/data/Voucher 43.pdf")
       

#if __name__=="__main__":
    # sg = matchSignature("C:/Users/Q0037/Downloads/signature_verification/data/Voucher 43.pdf")
    # print(sg)
    # if float(sg) > 0.2:
    #     print("True")
    #     signature_status = True
    # else:
    #     print("False")
    #     signature_status = False

